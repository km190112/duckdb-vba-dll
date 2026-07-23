//! クエリの実行と結果の組み立て。権限ゲート（層2）を必ず通す唯一の入口。

use crate::classify::{self, ApiKind};
use crate::conn::ConnState;
use crate::inbound::{self, InputGrid};
use crate::oleaut::VARIANT;
use crate::raw::Extracted;
use crate::value;
use crate::variant::Grid;

/// SELECT 系を実行し、1 行目をヘッダとする 2 次元 Variant 配列を返す。
pub fn query(state: &ConnState, sql: &str) -> Result<VARIANT, String> {
    classify::check(state.level, ApiKind::Query, &state.conn, sql)?;
    let mut result = state.conn.query(sql)?;
    result_to_grid(&mut result)?.into_variant()
}

/// 権限検査を伴わない内部専用のクエリ（`DackListTables` などの定型 SQL 用）。
///
/// 呼び出し側で SQL を組み立てているものにのみ使うこと。利用者入力は必ず [`query`] を通す。
pub(crate) fn query_internal(state: &ConnState, sql: &str) -> Result<VARIANT, String> {
    let mut result = state.conn.query(sql)?;
    result_to_grid(&mut result)?.into_variant()
}

/// DML を実行し、影響行数を返す。
pub fn execute(state: &ConnState, sql: &str) -> Result<i64, String> {
    classify::check(state.level, ApiKind::Execute, &state.conn, sql)?;
    let mut result = state.conn.query(sql)?;
    Ok(result.rows_changed())
}

/// DDL を実行する（管理者のみ）。影響行数の概念が無いので 0 を返す。
pub fn execute_ddl(state: &ConnState, sql: &str) -> Result<i64, String> {
    classify::check(state.level, ApiKind::ExecuteDdl, &state.conn, sql)?;
    let mut result = state.conn.query(sql)?;
    Ok(result.rows_changed())
}

/// 結果セットを [`Grid`] に変換する。
fn result_to_grid(result: &mut crate::raw::QueryResult) -> Result<Grid, String> {
    let ncols = result.column_count();
    if ncols == 0 {
        return Err(
            "この SQL は結果セットを返しません。DackExecute を使ってください。".to_string(),
        );
    }

    let headers: Vec<String> = (0..ncols).map(|c| result.column_name(c)).collect();

    // 行を 1 つも組み立てる前に、変換できない型が無いか検査する。
    // 途中まで貼ってから失敗する事態を避ける。
    for (c, name) in headers.iter().enumerate() {
        value::check_column_supported(name, result.column_type(c))?;
    }

    let mut grid = Grid::with_header(&headers);

    while let Some(chunk) = result.next_chunk() {
        let nrows = chunk.len();
        if nrows == 0 {
            continue;
        }
        // 列ごとにベクタを取り直すのは無駄なので、チャンク単位で先に集める。
        let vectors: Vec<_> = (0..ncols).map(|c| chunk.vector(c)).collect();

        for r in 0..nrows {
            for v in vectors.iter() {
                grid.push(unsafe { value::cell_to_variant(v, r) });
            }
        }

        // 1 チャンク積むごとに上限を確認する。1000 万行の SELECT で
        // メモリを食い潰してから失敗するのを防ぐ。
        grid.check_excel_limits()?;
    }

    Ok(grid)
}

// ---------------------------------------------------------------------------
// パラメータ付き実行（SQL インジェクション対策）
// ---------------------------------------------------------------------------

/// パラメータをバインドして SELECT を実行する。
pub fn query_params(state: &ConnState, sql: &str, params: &InputGrid) -> Result<VARIANT, String> {
    let prepared = prepare_checked(state, ApiKind::Query, sql)?;
    bind_all(&prepared, params)?;
    let mut result = prepared.execute()?;
    result_to_grid(&mut result)?.into_variant()
}

/// パラメータをバインドして DML を実行し、影響行数を返す。
pub fn execute_params(state: &ConnState, sql: &str, params: &InputGrid) -> Result<i64, String> {
    let prepared = prepare_checked(state, ApiKind::Execute, sql)?;
    bind_all(&prepared, params)?;
    let mut result = prepared.execute()?;
    Ok(result.rows_changed())
}

/// SQL を prepare し、権限ゲート（層2）を通す。
///
/// **パラメータ付き実行は 1 文のみ許可する。** 複数文を許すと、どの文に
/// どのパラメータが対応するのかが曖昧になり、検査と実行がずれる余地が生まれる。
fn prepare_checked(
    state: &ConnState,
    api: ApiKind,
    sql: &str,
) -> Result<crate::raw::Prepared, String> {
    let extracted = Extracted::extract(&state.conn, sql)?;
    if extracted.count() != 1 {
        return Err(format!(
            "パラメータ付きの実行では SQL を 1 文だけ指定してください（{} 文ありました）。",
            extracted.count()
        ));
    }
    let prepared = extracted.prepare(&state.conn, 0)?;
    classify::check_statement_type(state.level, api, prepared.statement_type())?;
    Ok(prepared)
}

/// 入力配列を `?` に順番にバインドする。
fn bind_all(prepared: &crate::raw::Prepared, params: &InputGrid) -> Result<(), String> {
    let values = flatten_params(params)?;
    let expected = prepared.param_count();

    if values.len() != expected {
        return Err(format!(
            "SQL の ? は {expected} 個ですが、パラメータが {} 個渡されました。",
            values.len()
        ));
    }

    for (i, (v, pos)) in values.iter().enumerate() {
        // value は次の反復まで生存する必要がある（bind はコピーを取る）。
        let value = unsafe { inbound::variant_to_value(v, pos) }?;
        unsafe { prepared.bind(i + 1, value.raw()) }?;
    }
    Ok(())
}

/// パラメータ配列を 1 次元に均す。
///
/// VBA の `Array(1, 2, 3)`（1 次元）、`Range("A1:C1").Value`（1 行 N 列）、
/// `Range("A1:A3").Value`（N 行 1 列）のいずれも受け付ける。
/// 縦横どちらで渡しても動くほうが、利用者が迷わないため。
fn flatten_params(params: &InputGrid) -> Result<Vec<(&crate::oleaut::VARIANT, String)>, String> {
    if params.rows == 1 {
        Ok((0..params.cols)
            .map(|c| (params.get(0, c), params.position(0, c)))
            .collect())
    } else if params.cols == 1 {
        Ok((0..params.rows)
            .map(|r| (params.get(r, 0), params.position(r, 0)))
            .collect())
    } else {
        Err(format!(
            "パラメータは 1 行または 1 列で渡してください（{} 行 {} 列が渡されました）。",
            params.rows, params.cols
        ))
    }
}

/// 1 行 1 列の結果を文字列として読む。スキーマ出力などの内部用。
///
/// 結果が空、または NULL の場合は `None`。
pub(crate) fn scalar_string(state: &ConnState, sql: &str) -> Result<Option<String>, String> {
    let mut result = state.conn.query(sql)?;
    if result.column_count() == 0 {
        return Ok(None);
    }
    let Some(chunk) = result.next_chunk() else {
        return Ok(None);
    };
    if chunk.is_empty() {
        return Ok(None);
    }
    let vec = chunk.vector(0);
    let mut v = unsafe { value::cell_to_variant(&vec, 0) };
    let out = if v.vt == crate::oleaut::VT_BSTR {
        Some(unsafe { crate::oleaut::bstr_to_string(v.value.bstrVal) })
    } else {
        None
    };
    v.clear();
    Ok(out)
}

/// `DackListTables`：テーブル一覧。
pub fn list_tables(state: &ConnState) -> Result<VARIANT, String> {
    query_internal(
        state,
        "SELECT table_schema AS スキーマ, table_name AS テーブル名, table_type AS 種別 \
         FROM information_schema.tables \
         ORDER BY table_schema, table_name",
    )
}

/// `DackDescribe`：列定義。テーブル名は識別子としてエスケープする。
pub fn describe(state: &ConnState, table: &str) -> Result<VARIANT, String> {
    let sql = format!(
        "SELECT column_name AS 列名, data_type AS 型, is_nullable AS NULL可, \
                column_default AS 既定値, ordinal_position AS 位置 \
         FROM information_schema.columns \
         WHERE table_name = {} \
         ORDER BY ordinal_position",
        quote_literal(table)
    );
    let v = query_internal(state, &sql)?;
    Ok(v)
}

/// SQL の文字列リテラルとして安全に引用する（シングルクォートを二重化）。
pub fn quote_literal(s: &str) -> String {
    format!("'{}'", s.replace('\'', "''"))
}

/// SQL の識別子として安全に引用する（ダブルクォートを二重化）。
pub fn quote_ident(s: &str) -> String {
    format!("\"{}\"", s.replace('"', "\"\""))
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::conn::{self, OpenOptions};
    use crate::level::Level;
    use crate::oleaut::*;

    fn tmp_db(name: &str) -> String {
        let dir = std::env::temp_dir().join("dackdb-tests");
        std::fs::create_dir_all(&dir).unwrap();
        let p = dir.join(name);
        let _ = std::fs::remove_file(&p);
        p.to_str().unwrap().to_string()
    }

    fn seeded(name: &str) -> String {
        let p = tmp_db(name);
        let h = conn::open(Level::Admin, &p, OpenOptions::default()).unwrap();
        conn::with_conn(h, |s| {
            s.conn
                .query("CREATE TABLE 売上 (id INTEGER PRIMARY KEY, 部署 VARCHAR, 金額 BIGINT, 日付 DATE)")
                .unwrap();
            s.conn
                .query(
                    "INSERT INTO 売上 VALUES \
                     (1, '営業部', 1000, DATE '2024-01-15'), \
                     (2, '技術部𠮷😀', 2000, DATE '2024-02-20'), \
                     (3, NULL, NULL, NULL)",
                )
                .unwrap();
        })
        .unwrap();
        conn::close(h).unwrap();
        p
    }

    unsafe fn cell(psa: *mut SAFEARRAY, row: i32, col: i32) -> VARIANT {
        let mut out = VARIANT::empty();
        VariantInit(&mut out);
        let idx = [row, col];
        assert_eq!(
            SafeArrayGetElement(psa, idx.as_ptr(), &mut out as *mut VARIANT as *mut _),
            S_OK
        );
        out
    }

    #[test]
    fn query_returns_header_plus_rows_with_japanese_intact() {
        let p = seeded("query.db");
        let h = conn::open(Level::Read, &p, OpenOptions::default()).unwrap();
        let mut v = conn::with_conn(h, |s| query(s, "SELECT * FROM 売上 ORDER BY id"))
            .unwrap()
            .unwrap();

        unsafe {
            let psa = v.value.parray;
            let mut ub_r = 0i32;
            let mut ub_c = 0i32;
            SafeArrayGetUBound(psa, 1, &mut ub_r);
            SafeArrayGetUBound(psa, 2, &mut ub_c);
            assert_eq!((ub_r, ub_c), (4, 4), "ヘッダ1行 + データ3行, 4列");

            // ヘッダの日本語列名
            let mut hdr = cell(psa, 1, 2);
            assert_eq!(bstr_to_string(hdr.value.bstrVal), "部署");
            hdr.clear();

            // サロゲートペア・絵文字を含むデータ
            let mut name = cell(psa, 3, 2);
            assert_eq!(bstr_to_string(name.value.bstrVal), "技術部𠮷😀");
            name.clear();

            // NULL は VT_NULL（VBA の IsNull() が True、セルは空欄）
            assert_eq!(cell(psa, 4, 2).vt, VT_NULL);
            assert_eq!(cell(psa, 4, 3).vt, VT_NULL);

            // BIGINT は VT_I8
            assert_eq!(cell(psa, 2, 3).vt, VT_I8);
            assert_eq!(cell(psa, 2, 3).value.llVal, 1000);

            // DATE は VT_DATE。2024-01-15 の Excel シリアル値は 45306
            let d = cell(psa, 2, 4);
            assert_eq!(d.vt, VT_DATE);
            assert_eq!(d.value.date, 45306.0);
        }
        v.clear();
        conn::close(h).unwrap();
    }

    #[test]
    fn empty_result_returns_header_only() {
        let p = seeded("empty.db");
        let h = conn::open(Level::Read, &p, OpenOptions::default()).unwrap();
        let mut v = conn::with_conn(h, |s| query(s, "SELECT * FROM 売上 WHERE id = 999"))
            .unwrap()
            .unwrap();
        unsafe {
            let mut ub_r = 0i32;
            SafeArrayGetUBound(v.value.parray, 1, &mut ub_r);
            assert_eq!(ub_r, 1, "ヘッダ行だけが返るはず");
        }
        v.clear();
        conn::close(h).unwrap();
    }

    // ---- 権限ゲートのバイパス試験 ----

    #[test]
    fn multi_statement_smuggling_is_blocked() {
        let p = seeded("smuggle.db");
        let h = conn::open(Level::ReadWrite, &p, OpenOptions::default()).unwrap();
        conn::with_conn(h, |s| {
            // 後半に DDL を紛れ込ませる古典的な手口
            let err = execute(
                s,
                "INSERT INTO 売上 VALUES (9,'x',1,NULL); DROP TABLE 売上;",
            )
            .unwrap_err();
            assert!(err.contains("DROP"), "DROP が検出されていない: {err}");

            // 拒否されたなら INSERT も実行されていないこと（部分実行の禁止）
            let mut r = s.conn.query("SELECT count(*) FROM 売上").unwrap();
            let chunk = r.next_chunk().unwrap();
            let v = chunk.vector(0);
            let n = unsafe { *(v.data_ptr() as *const i64) };
            assert_eq!(n, 3, "拒否されたのに INSERT が実行されている");
        })
        .unwrap();
        conn::close(h).unwrap();
    }

    #[test]
    fn cte_wrapped_insert_is_detected_as_insert_not_select() {
        let p = seeded("cte.db");
        let h = conn::open(Level::Read, &p, OpenOptions::default()).unwrap();
        conn::with_conn(h, |s| {
            // 先頭キーワードは WITH だが実体は INSERT。正規表現方式なら通ってしまう。
            let err = query(
                s,
                "WITH x AS (SELECT 9 AS id) INSERT INTO 売上 SELECT id,'z',1,NULL FROM x",
            )
            .unwrap_err();
            assert!(!err.is_empty(), "CTE で包んだ INSERT が通ってしまった");
        })
        .unwrap();
        conn::close(h).unwrap();
    }

    #[test]
    fn comment_disguised_ddl_is_blocked() {
        let p = seeded("comment.db");
        let h = conn::open(Level::ReadWrite, &p, OpenOptions::default()).unwrap();
        conn::with_conn(h, |s| {
            let err = execute(s, "/* SELECT 1 */ DROP TABLE 売上").unwrap_err();
            assert!(err.contains("DROP"), "{err}");
        })
        .unwrap();
        conn::close(h).unwrap();
    }

    #[test]
    fn dack_query_refuses_writes_even_at_admin_level() {
        let p = seeded("qwrite.db");
        let h = conn::open(Level::Admin, &p, OpenOptions::default()).unwrap();
        conn::with_conn(h, |s| {
            let err = query(s, "INSERT INTO 売上 VALUES (9,'x',1,NULL)").unwrap_err();
            assert!(err.contains("DackQuery"), "{err}");
            assert!(err.contains("DackExecute"), "代替 API の案内が無い: {err}");
        })
        .unwrap();
        conn::close(h).unwrap();
    }

    #[test]
    fn readwrite_cannot_create_or_drop() {
        let p = seeded("rwddl.db");
        let h = conn::open(Level::ReadWrite, &p, OpenOptions::default()).unwrap();
        conn::with_conn(h, |s| {
            for sql in [
                "CREATE TABLE t2 (a INTEGER)",
                "DROP TABLE 売上",
                "ALTER TABLE 売上 ADD COLUMN 備考 VARCHAR",
                "ATTACH 'other.db' AS o",
                "SET access_mode='READ_WRITE'",
                "INSTALL httpfs",
            ] {
                let r = execute_ddl(s, sql);
                let err = match r {
                    Ok(_) => panic!("「{sql}」が読み書き DLL で実行できてしまった"),
                    Err(e) => e,
                };
                assert!(
                    err.contains("dackdb_admin.dll"),
                    "「{sql}」が管理者 DLL への案内無しに扱われた: {err}"
                );
            }
        })
        .unwrap();
        conn::close(h).unwrap();
    }

    #[test]
    fn read_level_cannot_insert_through_execute() {
        let p = seeded("rexec.db");
        let h = conn::open(Level::Read, &p, OpenOptions::default()).unwrap();
        conn::with_conn(h, |s| {
            let err = execute(s, "INSERT INTO 売上 VALUES (9,'x',1,NULL)").unwrap_err();
            assert!(
                err.contains("dackdb_rw.dll"),
                "上位 DLL の案内が無い: {err}"
            );
        })
        .unwrap();
        conn::close(h).unwrap();
    }

    #[test]
    fn admin_can_do_ddl() {
        let p = seeded("admin.db");
        let h = conn::open(Level::Admin, &p, OpenOptions::default()).unwrap();
        conn::with_conn(h, |s| {
            execute_ddl(s, "CREATE TABLE 新表 (a INTEGER PRIMARY KEY, b VARCHAR)").unwrap();
            execute_ddl(s, "ALTER TABLE 新表 ADD COLUMN c DATE").unwrap();
            execute(s, "INSERT INTO 新表 VALUES (1,'あ',NULL)").unwrap();
            execute_ddl(s, "DROP TABLE 新表").unwrap();
        })
        .unwrap();
        conn::close(h).unwrap();
    }

    #[test]
    fn syntax_errors_report_duckdb_message() {
        let p = seeded("syntax.db");
        let h = conn::open(Level::Read, &p, OpenOptions::default()).unwrap();
        conn::with_conn(h, |s| {
            let err = query(s, "SELEC * FROM 売上").unwrap_err();
            assert!(err.contains("構文エラー"), "{err}");
        })
        .unwrap();
        conn::close(h).unwrap();
    }

    // ---- パラメータバインド ----

    /// VBA の Array(...) / Range(...).Value を模した 1 行 N 列の配列を作る。
    fn params_of(values: Vec<VARIANT>) -> VARIANT {
        let mut g = Grid::with_cols(values.len());
        g.push_row(values);
        g.into_variant().unwrap()
    }

    #[test]
    fn query_params_binds_values_including_japanese() {
        let p = seeded("qparams.db");
        let h = conn::open(Level::Read, &p, OpenOptions::default()).unwrap();
        let mut input = params_of(vec![VARIANT::bstr("技術部𠮷😀")]);

        let mut v = conn::with_conn(h, |s| {
            let g = unsafe { inbound::read_input_grid(&input, "パラメータ") }.unwrap();
            query_params(s, "SELECT id FROM 売上 WHERE 部署 = ?", &g)
        })
        .unwrap()
        .unwrap();

        unsafe {
            let mut ub = 0i32;
            SafeArrayGetUBound(v.value.parray, 1, &mut ub);
            assert_eq!(ub, 2, "ヘッダ + 1 件ヒットするはず");
            assert_eq!(cell(v.value.parray, 2, 1).value.lVal, 2);
        }
        v.clear();
        input.clear();
        conn::close(h).unwrap();
    }

    /// パラメータは値として扱われ、SQL として解釈されないこと。
    /// これがバインドを用意した理由そのもの。
    #[test]
    fn injection_payload_in_a_parameter_is_treated_as_data() {
        let p = seeded("inject.db");
        let h = conn::open(Level::ReadWrite, &p, OpenOptions::default()).unwrap();
        let mut input = params_of(vec![VARIANT::bstr("' OR 1=1; DROP TABLE 売上; --")]);

        conn::with_conn(h, |s| {
            let g = unsafe { inbound::read_input_grid(&input, "パラメータ") }.unwrap();
            // 実行は成功する。ただしヒット 0 件で、テーブルも無事であること。
            let mut v = query_params(s, "SELECT id FROM 売上 WHERE 部署 = ?", &g).unwrap();
            unsafe {
                let mut ub = 0i32;
                SafeArrayGetUBound(v.value.parray, 1, &mut ub);
                assert_eq!(ub, 1, "ヘッダのみ（0 件ヒット）のはず");
            }
            v.clear();

            // テーブルが消えていないこと
            s.conn.query("SELECT count(*) FROM 売上").unwrap();
        })
        .unwrap();

        input.clear();
        conn::close(h).unwrap();
    }

    #[test]
    fn parameter_count_mismatch_is_reported() {
        let p = seeded("pcount.db");
        let h = conn::open(Level::Read, &p, OpenOptions::default()).unwrap();
        let mut input = params_of(vec![VARIANT::i32(1), VARIANT::i32(2)]);

        conn::with_conn(h, |s| {
            let g = unsafe { inbound::read_input_grid(&input, "パラメータ") }.unwrap();
            let err = query_params(s, "SELECT * FROM 売上 WHERE id = ?", &g).unwrap_err();
            assert!(err.contains("1 個"), "SQL 側の個数が示されていない: {err}");
            assert!(err.contains("2 個"), "渡された個数が示されていない: {err}");
        })
        .unwrap();

        input.clear();
        conn::close(h).unwrap();
    }

    /// パラメータ付き実行でも権限ゲートが効くこと。
    /// prepared 経由は classify::check を通らない別経路なので、独立に確認する。
    #[test]
    fn params_path_still_enforces_permissions() {
        let p = seeded("pperm.db");
        let h = conn::open(Level::Read, &p, OpenOptions::default()).unwrap();
        let mut input = params_of(vec![VARIANT::i32(9)]);

        conn::with_conn(h, |s| {
            let g = unsafe { inbound::read_input_grid(&input, "パラメータ") }.unwrap();
            let err = execute_params(s, "DELETE FROM 売上 WHERE id = ?", &g).unwrap_err();
            assert!(
                err.contains("dackdb_rw.dll"),
                "権限ゲートが効いていない: {err}"
            );
        })
        .unwrap();

        input.clear();
        conn::close(h).unwrap();
    }

    /// パラメータ付き実行は 1 文のみ。複数文だと検査と実行がずれる余地が出る。
    #[test]
    fn params_path_rejects_multiple_statements() {
        let p = seeded("pmulti.db");
        let h = conn::open(Level::ReadWrite, &p, OpenOptions::default()).unwrap();
        let mut input = params_of(vec![VARIANT::i32(1)]);

        conn::with_conn(h, |s| {
            let g = unsafe { inbound::read_input_grid(&input, "パラメータ") }.unwrap();
            let err = execute_params(s, "DELETE FROM 売上 WHERE id = ?; DROP TABLE 売上;", &g)
                .unwrap_err();
            assert!(err.contains("1 文だけ"), "{err}");
        })
        .unwrap();

        input.clear();
        conn::close(h).unwrap();
    }

    /// 縦横どちらの向きで渡しても同じように動くこと。
    #[test]
    fn params_accept_both_row_and_column_orientation() {
        let p = seeded("porient.db");
        let h = conn::open(Level::Read, &p, OpenOptions::default()).unwrap();

        // 1 行 2 列（Range("A1:B1").Value 相当）
        let mut horizontal = params_of(vec![VARIANT::i32(1), VARIANT::i32(2)]);
        // 2 行 1 列（Range("A1:A2").Value 相当）
        let mut g2 = Grid::with_cols(1);
        g2.push_row(vec![VARIANT::i32(1)]);
        g2.push_row(vec![VARIANT::i32(2)]);
        let mut vertical = g2.into_variant().unwrap();

        conn::with_conn(h, |s| {
            for (input, label) in [(&horizontal, "1 行 2 列"), (&vertical, "2 行 1 列")] {
                let g = unsafe { inbound::read_input_grid(input, "パラメータ") }.unwrap();
                let mut v =
                    query_params(s, "SELECT id FROM 売上 WHERE id IN (?, ?) ORDER BY id", &g)
                        .unwrap_or_else(|e| panic!("{label} で失敗: {e}"));
                unsafe {
                    let mut ub = 0i32;
                    SafeArrayGetUBound(v.value.parray, 1, &mut ub);
                    assert_eq!(ub, 3, "{label}: ヘッダ + 2 件");
                }
                v.clear();
            }
        })
        .unwrap();

        horizontal.clear();
        vertical.clear();
        conn::close(h).unwrap();
    }

    #[test]
    fn quoting_helpers_escape_correctly() {
        assert_eq!(quote_literal("O'Brien"), "'O''Brien'");
        assert_eq!(quote_ident("my\"col"), "\"my\"\"col\"");
        assert_eq!(quote_literal("売上"), "'売上'");
    }

    #[test]
    fn describe_returns_column_metadata() {
        let p = seeded("desc.db");
        let h = conn::open(Level::Read, &p, OpenOptions::default()).unwrap();
        let mut v = conn::with_conn(h, |s| describe(s, "売上"))
            .unwrap()
            .unwrap();
        unsafe {
            let mut ub_r = 0i32;
            SafeArrayGetUBound(v.value.parray, 1, &mut ub_r);
            assert_eq!(ub_r, 5, "ヘッダ + 4 列");
        }
        v.clear();
        conn::close(h).unwrap();
    }

    #[test]
    fn list_tables_includes_seeded_table() {
        let p = seeded("list.db");
        let h = conn::open(Level::Read, &p, OpenOptions::default()).unwrap();
        let mut v = conn::with_conn(h, list_tables).unwrap().unwrap();
        unsafe {
            let mut name = cell(v.value.parray, 2, 2);
            assert_eq!(bstr_to_string(name.value.bstrVal), "売上");
            name.clear();
        }
        v.clear();
        conn::close(h).unwrap();
    }
}
