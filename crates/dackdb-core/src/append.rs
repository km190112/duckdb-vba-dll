//! シート範囲の一括投入（`DackAppendArray`）。
//!
//! DuckDB の Appender API を使う。INSERT 文を 1 行ずつ実行するより 1〜2 桁速く、
//! Excel → DB 方向の主力経路になる。

use crate::conn::ConnState;
use crate::inbound::{self, InputGrid};
use crate::raw::Appender;

/// 2 次元配列をテーブルへ一括投入し、投入行数を返す。
///
/// # 列の対応
///
/// 配列の列は**テーブルの列定義の順に**対応する（位置指定）。ヘッダ行は扱わないので、
/// 見出しを含まない範囲（例: `Range("A2:F1000").Value`）を渡すこと。
///
/// # 全か無か
///
/// 途中で失敗しても中途半端に行が入らないよう、全体をトランザクションで囲む。
/// Appender は内部バッファが埋まると自動的に書き出すため、トランザクションが無いと
/// 「1000 行目でエラー、でも 800 行は入っている」という状態が起こり得る。
pub fn append_array(state: &ConnState, table: &str, data: &InputGrid) -> Result<i64, String> {
    let appender = Appender::new(&state.conn, table)?;

    let table_cols = appender.column_count();
    if table_cols != data.cols {
        return Err(format!(
            "テーブル「{table}」は {table_cols} 列ですが、渡された範囲は {} 列です。\n\
             列数を合わせてください（見出し行は含めず、テーブルの列順に並べます）。",
            data.cols
        ));
    }

    state.conn.query("BEGIN TRANSACTION")?;

    match fill(&appender, data) {
        Ok(()) => {
            // flush を明示的に呼んで失敗を検出する。Drop 任せだと握り潰される。
            if let Err(e) = appender.flush() {
                let _ = state.conn.query("ROLLBACK");
                return Err(e);
            }
            drop(appender);
            state.conn.query("COMMIT")?;
            Ok(data.rows as i64)
        }
        Err(e) => {
            drop(appender);
            let _ = state.conn.query("ROLLBACK");
            Err(e)
        }
    }
}

fn fill(appender: &Appender, data: &InputGrid) -> Result<(), String> {
    for r in 0..data.rows {
        for c in 0..data.cols {
            let pos = data.position(r, c);
            let value = unsafe { inbound::variant_to_value(data.get(r, c), &pos) }?;
            unsafe { appender.append_value(value.raw()) }.map_err(|e| format!("{pos}: {e}"))?;
        }
        appender
            .end_row()
            .map_err(|e| format!("{} 行目: {e}", r + 1))?;
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::conn::{self, OpenOptions};
    use crate::level::Level;
    use crate::oleaut::*;
    use crate::variant::Grid;

    fn seeded(name: &str) -> String {
        let dir = std::env::temp_dir().join("dackdb-tests");
        std::fs::create_dir_all(&dir).unwrap();
        let p = dir.join(name);
        let _ = std::fs::remove_file(&p);
        let path = p.to_str().unwrap().to_string();

        let h = conn::open(Level::Admin, &path, OpenOptions::default()).unwrap();
        conn::with_conn(h, |s| {
            s.conn
                .query("CREATE TABLE 売上 (id INTEGER, 部署 VARCHAR, 金額 BIGINT, 日付 DATE)")
                .unwrap();
        })
        .unwrap();
        conn::close(h).unwrap();
        path
    }

    /// VBA から渡される 2 次元 Variant 配列を模して作る。
    fn make_input(rows: Vec<Vec<VARIANT>>) -> VARIANT {
        let cols = rows[0].len();
        let mut g = Grid::with_cols(cols);
        for r in rows {
            g.push_row(r);
        }
        g.into_variant().unwrap()
    }

    fn count_rows(state: &ConnState) -> i64 {
        let mut r = state.conn.query("SELECT count(*) FROM 売上").unwrap();
        let chunk = r.next_chunk().unwrap();
        let v = chunk.vector(0);
        unsafe { *(v.data_ptr() as *const i64) }
    }

    #[test]
    fn appends_rows_with_japanese_and_nulls() {
        let path = seeded("append_ok.db");
        let h = conn::open(Level::Admin, &path, OpenOptions::default()).unwrap();

        let mut input = make_input(vec![
            vec![
                VARIANT::i32(1),
                VARIANT::bstr("技術部𠮷😀"),
                VARIANT::i64(1000),
                VARIANT::date(45306.0), // 2024-01-15
            ],
            vec![
                VARIANT::i32(2),
                VARIANT::null(),
                VARIANT::empty(), // 空セル → NULL
                VARIANT::null(),
            ],
        ]);

        conn::with_conn(h, |s| {
            let grid = unsafe { inbound::read_input_grid(&input, "データ") }.unwrap();
            let n = append_array(s, "売上", &grid).unwrap();
            assert_eq!(n, 2);
            assert_eq!(count_rows(s), 2);

            // 日本語が壊れていないこと
            let mut r = s.conn.query("SELECT 部署 FROM 売上 WHERE id = 1").unwrap();
            let chunk = r.next_chunk().unwrap();
            let v = chunk.vector(0);
            let mut cell = unsafe { crate::value::cell_to_variant(&v, 0) };
            assert_eq!(unsafe { bstr_to_string(cell.value.bstrVal) }, "技術部𠮷😀");
            cell.clear();

            // 日付が正しく入っていること
            let mut r = s
                .conn
                .query("SELECT 日付::VARCHAR FROM 売上 WHERE id = 1")
                .unwrap();
            let chunk = r.next_chunk().unwrap();
            let v = chunk.vector(0);
            let mut cell = unsafe { crate::value::cell_to_variant(&v, 0) };
            assert_eq!(unsafe { bstr_to_string(cell.value.bstrVal) }, "2024-01-15");
            cell.clear();

            // NULL が NULL として入っていること
            let mut r = s
                .conn
                .query("SELECT count(*) FROM 売上 WHERE 部署 IS NULL AND 金額 IS NULL")
                .unwrap();
            let chunk = r.next_chunk().unwrap();
            let v = chunk.vector(0);
            assert_eq!(unsafe { *(v.data_ptr() as *const i64) }, 1);
        })
        .unwrap();

        input.clear();
        conn::close(h).unwrap();
    }

    #[test]
    fn column_count_mismatch_is_reported_clearly() {
        let path = seeded("append_cols.db");
        let h = conn::open(Level::Admin, &path, OpenOptions::default()).unwrap();
        let mut input = make_input(vec![vec![VARIANT::i32(1), VARIANT::bstr("営業部")]]);

        conn::with_conn(h, |s| {
            let grid = unsafe { inbound::read_input_grid(&input, "データ") }.unwrap();
            let err = append_array(s, "売上", &grid).unwrap_err();
            assert!(
                err.contains("4 列"),
                "テーブルの列数が示されていない: {err}"
            );
            assert!(err.contains("2 列"), "渡された列数が示されていない: {err}");
        })
        .unwrap();

        input.clear();
        conn::close(h).unwrap();
    }

    /// 途中の行でエラーが起きても 1 行も入らないこと。
    /// Appender は内部バッファが埋まると自動で書き出すため、
    /// トランザクションが無いと部分的に入ってしまう。
    #[test]
    fn a_failing_row_rolls_back_the_whole_batch() {
        let path = seeded("append_rollback.db");
        let h = conn::open(Level::Admin, &path, OpenOptions::default()).unwrap();

        // 3000 行目に Excel のエラー値を混ぜる（Appender の内部バッファ 2048 行を超える位置）
        let mut rows = Vec::new();
        for i in 0..4000i32 {
            let third = if i == 3000 {
                let mut e = VARIANT::empty();
                e.vt = VT_ERROR;
                e.value.lVal = -2146826246; // xlErrNA
                e
            } else {
                VARIANT::i64(i as i64 * 100)
            };
            rows.push(vec![
                VARIANT::i32(i),
                VARIANT::bstr("営業部"),
                third,
                VARIANT::null(),
            ]);
        }
        let mut input = make_input(rows);

        conn::with_conn(h, |s| {
            let grid = unsafe { inbound::read_input_grid(&input, "データ") }.unwrap();
            let err = append_array(s, "売上", &grid).unwrap_err();
            assert!(err.contains("3001 行"), "失敗位置が示されていない: {err}");
            assert_eq!(count_rows(s), 0, "失敗したのに行が残っている（部分投入）");
        })
        .unwrap();

        input.clear();
        conn::close(h).unwrap();
    }

    #[test]
    fn unknown_table_is_reported_clearly() {
        let path = seeded("append_notable.db");
        let h = conn::open(Level::Admin, &path, OpenOptions::default()).unwrap();
        let mut input = make_input(vec![vec![VARIANT::i32(1)]]);

        conn::with_conn(h, |s| {
            let grid = unsafe { inbound::read_input_grid(&input, "データ") }.unwrap();
            let err = append_array(s, "存在しない表", &grid).unwrap_err();
            assert!(!err.is_empty());
        })
        .unwrap();

        input.clear();
        conn::close(h).unwrap();
    }

    #[test]
    fn large_batch_is_fast_and_complete() {
        let path = seeded("append_perf.db");
        let h = conn::open(Level::Admin, &path, OpenOptions::default()).unwrap();

        let mut rows = Vec::with_capacity(50_000);
        for i in 0..50_000i32 {
            rows.push(vec![
                VARIANT::i32(i),
                VARIANT::bstr("営業部"),
                VARIANT::i64(i as i64),
                VARIANT::null(),
            ]);
        }
        let mut input = make_input(rows);

        conn::with_conn(h, |s| {
            let grid = unsafe { inbound::read_input_grid(&input, "データ") }.unwrap();
            let n = append_array(s, "売上", &grid).unwrap();
            assert_eq!(n, 50_000);
            assert_eq!(count_rows(s), 50_000);
        })
        .unwrap();

        input.clear();
        conn::close(h).unwrap();
    }
}
