//! スキーマ情報の出力（管理者 DLL のみ）。

use crate::conn::ConnState;
use crate::oleaut::VARIANT;
use crate::query;

/// PRIMARY KEY 情報を含む列定義の一覧を返す SQL。
///
/// `duckdb_constraints()` の `constraint_column_names` は LIST 型なので、
/// そのままでは Excel のセルに変換できない。`list_contains` で判定して
/// VARCHAR の「PK」フラグに落とし込んでいる。
const SCHEMA_TABLE_SQL: &str = r#"
SELECT c.table_schema                        AS "スキーマ",
       c.table_name                          AS "テーブル",
       c.ordinal_position                    AS "位置",
       c.column_name                         AS "列名",
       c.data_type                           AS "型",
       c.is_nullable                         AS "NULL可",
       COALESCE(c.column_default, '')        AS "既定値",
       CASE WHEN pk.cols IS NOT NULL AND list_contains(pk.cols, c.column_name)
            THEN 'PK' ELSE '' END            AS "キー"
FROM information_schema.columns c
LEFT JOIN (
    SELECT schema_name, table_name, constraint_column_names AS cols
    FROM duckdb_constraints()
    WHERE constraint_type = 'PRIMARY KEY'
) pk
  ON pk.table_name = c.table_name AND pk.schema_name = c.table_schema
ORDER BY c.table_schema, c.table_name, c.ordinal_position
"#;

/// `duckdb_tables()` の `sql` 列に各テーブルの CREATE 文が入っている。
const SCHEMA_DDL_SQL: &str = r#"
SELECT string_agg(sql, chr(10) || chr(10) ORDER BY schema_name, table_name)
FROM duckdb_tables()
"#;

/// `DackExportSchema` の実装。
///
/// - `"table"`（既定）: 2 次元配列。そのままシートに貼れる。
/// - `"ddl"`: CREATE 文を連結した 1 本の文字列。
pub fn export(state: &ConnState, format: &str) -> Result<VARIANT, String> {
    match format.trim().to_ascii_lowercase().as_str() {
        "" | "table" => query::query_internal(state, SCHEMA_TABLE_SQL),
        "ddl" => {
            let ddl = query::scalar_string(state, SCHEMA_DDL_SQL)?;
            match ddl {
                Some(s) => Ok(VARIANT::bstr(&s)),
                None => Ok(VARIANT::bstr("-- テーブルがありません")),
            }
        }
        other => Err(format!(
            "未知の形式「{other}」です。\"table\"（2 次元配列）か \"ddl\"（CREATE 文）を指定してください。"
        )),
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::conn::{self, OpenOptions};
    use crate::level::Level;
    use crate::oleaut::*;

    fn seeded(name: &str) -> String {
        let dir = std::env::temp_dir().join("dackdb-tests");
        std::fs::create_dir_all(&dir).unwrap();
        let p = dir.join(name);
        let _ = std::fs::remove_file(&p);
        let path = p.to_str().unwrap().to_string();

        let h = conn::open(Level::Admin, &path, OpenOptions::default()).unwrap();
        conn::with_conn(h, |s| {
            s.conn
                .query(
                    "CREATE TABLE 社員 (\
                       社員番号 INTEGER PRIMARY KEY, \
                       氏名 VARCHAR NOT NULL, \
                       入社日 DATE DEFAULT '2000-01-01')",
                )
                .unwrap();
        })
        .unwrap();
        conn::close(h).unwrap();
        path
    }

    #[test]
    fn table_format_marks_primary_key_and_keeps_japanese() {
        let path = seeded("schema_table.db");
        let h = conn::open(Level::Admin, &path, OpenOptions::default()).unwrap();
        let mut v = conn::with_conn(h, |s| export(s, "table")).unwrap().unwrap();

        unsafe {
            let psa = v.value.parray;
            let mut ub_r = 0i32;
            let mut ub_c = 0i32;
            SafeArrayGetUBound(psa, 1, &mut ub_r);
            SafeArrayGetUBound(psa, 2, &mut ub_c);
            assert_eq!(ub_r, 4, "ヘッダ + 3 列分の行");
            assert_eq!(ub_c, 8, "8 項目");

            let get = |r: i32, c: i32| -> String {
                let mut out = VARIANT::empty();
                VariantInit(&mut out);
                let idx = [r, c];
                SafeArrayGetElement(psa, idx.as_ptr(), &mut out as *mut VARIANT as *mut _);
                let s = if out.vt == VT_BSTR {
                    bstr_to_string(out.value.bstrVal)
                } else {
                    String::new()
                };
                out.clear();
                s
            };

            assert_eq!(get(1, 8), "キー", "ヘッダの日本語");
            assert_eq!(get(2, 4), "社員番号", "1 列目の列名");
            assert_eq!(get(2, 8), "PK", "主キーが PK と印されていない");
            assert_eq!(get(3, 4), "氏名");
            assert_eq!(get(3, 8), "", "非主キー列に PK が付いている");
        }
        v.clear();
        conn::close(h).unwrap();
    }

    #[test]
    fn ddl_format_returns_create_statements() {
        let path = seeded("schema_ddl.db");
        let h = conn::open(Level::Admin, &path, OpenOptions::default()).unwrap();
        let mut v = conn::with_conn(h, |s| export(s, "ddl")).unwrap().unwrap();
        assert_eq!(v.vt, VT_BSTR);
        let ddl = unsafe { bstr_to_string(v.value.bstrVal) };
        assert!(ddl.contains("CREATE TABLE"), "CREATE 文が無い: {ddl}");
        assert!(ddl.contains("社員"), "テーブル名が無い: {ddl}");
        v.clear();
        conn::close(h).unwrap();
    }

    #[test]
    fn unknown_format_is_rejected_with_valid_options() {
        let path = seeded("schema_bad.db");
        let h = conn::open(Level::Admin, &path, OpenOptions::default()).unwrap();
        let err = conn::with_conn(h, |s| export(s, "json")).unwrap().unwrap_err();
        assert!(err.contains("table"), "{err}");
        assert!(err.contains("ddl"), "{err}");
        conn::close(h).unwrap();
    }

    #[test]
    fn empty_format_defaults_to_table() {
        let path = seeded("schema_default.db");
        let h = conn::open(Level::Admin, &path, OpenOptions::default()).unwrap();
        let mut v = conn::with_conn(h, |s| export(s, "")).unwrap().unwrap();
        assert_eq!(v.vt, VT_ARRAY | VT_VARIANT);
        v.clear();
        conn::close(h).unwrap();
    }
}
