//! 管理者 DLL で使えるデータベース機能の確認。
//!
//! マニュアルに書いている内容が実際に動くことを固定する。
//! ここが落ちたら、マニュアルの記述も直す必要がある。

use dackdb_core::conn::{self, OpenOptions};
use dackdb_core::level::Level;
use dackdb_core::query;

fn fresh_db(name: &str) -> i64 {
    let dir = std::env::temp_dir().join("dackdb-ddl-tests");
    std::fs::create_dir_all(&dir).unwrap();
    let p = dir.join(name);
    let _ = std::fs::remove_file(&p);
    conn::open(Level::Admin, p.to_str().unwrap(), OpenOptions::default()).unwrap()
}

/// DDL を実行する（成功を期待）。
fn ddl(h: i64, sql: &str) {
    conn::with_conn(h, |s| query::execute_ddl(s, sql))
        .unwrap()
        .unwrap_or_else(|e| panic!("失敗した DDL: {sql}\n  {e}"));
}

/// DML を実行する（成功を期待）。
fn dml(h: i64, sql: &str) -> i64 {
    conn::with_conn(h, |s| query::execute(s, sql))
        .unwrap()
        .unwrap_or_else(|e| panic!("失敗した DML: {sql}\n  {e}"))
}

/// DML が拒否されることを期待し、エラーメッセージを返す。
fn dml_err(h: i64, sql: &str) -> String {
    match conn::with_conn(h, |s| query::execute(s, sql)).unwrap() {
        Ok(_) => panic!("拒否されるべき DML が通ってしまった: {sql}"),
        Err(e) => e,
    }
}

fn ddl_err(h: i64, sql: &str) -> String {
    match conn::with_conn(h, |s| query::execute_ddl(s, sql)).unwrap() {
        Ok(_) => panic!("拒否されるべき DDL が通ってしまった: {sql}"),
        Err(e) => e,
    }
}

// ---------------------------------------------------------------------------
// 制約
// ---------------------------------------------------------------------------

#[test]
fn all_constraint_kinds_can_be_created_and_are_enforced() {
    let h = fresh_db("constraints.db");

    ddl(
        h,
        "CREATE TABLE 部署 (部署コード INTEGER PRIMARY KEY, 部署名 VARCHAR NOT NULL)",
    );
    ddl(
        h,
        "CREATE TABLE 社員 (
           社員番号 INTEGER PRIMARY KEY,
           氏名     VARCHAR NOT NULL,
           メール   VARCHAR UNIQUE,
           年齢     INTEGER CHECK (年齢 >= 0),
           区分     VARCHAR DEFAULT '正社員',
           部署コード INTEGER REFERENCES 部署(部署コード)
         )",
    );

    dml(h, "INSERT INTO 部署 VALUES (1, '営業部')");
    dml(
        h,
        "INSERT INTO 社員 (社員番号, 氏名, メール, 年齢, 部署コード) VALUES (1,'山田','a@x.jp',30,1)",
    );

    // 既定値が効いていること
    let n = conn::with_conn(h, |s| {
        query::execute(s, "UPDATE 社員 SET 氏名 = 氏名 WHERE 区分 = '正社員'")
    })
    .unwrap()
    .unwrap();
    assert_eq!(n, 1, "DEFAULT が適用されていない");

    // 各制約が実際に違反を弾くこと
    assert!(dml_err(h, "INSERT INTO 部署 VALUES (1,'重複')").contains("primary key"));
    assert!(dml_err(h, "INSERT INTO 部署 VALUES (2, NULL)").contains("NOT NULL"));
    assert!(dml_err(
        h,
        "INSERT INTO 社員 (社員番号,氏名,メール) VALUES (2,'鈴木','a@x.jp')"
    )
    .to_lowercase()
    .contains("constraint"));
    assert!(dml_err(
        h,
        "INSERT INTO 社員 (社員番号,氏名,年齢) VALUES (3,'佐藤',-1)"
    )
    .to_lowercase()
    .contains("constraint"));
    assert!(dml_err(
        h,
        "INSERT INTO 社員 (社員番号,氏名,部署コード) VALUES (4,'高橋',999)"
    )
    .contains("foreign key"));

    conn::close(h).unwrap();
}

#[test]
fn composite_primary_key_works() {
    let h = fresh_db("composite_pk.db");
    ddl(
        h,
        "CREATE TABLE 勤怠 (社員番号 INTEGER, 日付 DATE, 時間 DOUBLE, PRIMARY KEY (社員番号, 日付))",
    );
    dml(h, "INSERT INTO 勤怠 VALUES (1, DATE '2024-01-15', 8.0)");
    dml(h, "INSERT INTO 勤怠 VALUES (1, DATE '2024-01-16', 7.5)");
    assert!(
        dml_err(h, "INSERT INTO 勤怠 VALUES (1, DATE '2024-01-15', 9.0)").contains("primary key")
    );
    conn::close(h).unwrap();
}

#[test]
fn upsert_on_conflict_works() {
    let h = fresh_db("upsert.db");
    ddl(
        h,
        "CREATE TABLE 在庫 (品番 VARCHAR PRIMARY KEY, 数量 INTEGER)",
    );
    dml(h, "INSERT INTO 在庫 VALUES ('A', 10)");
    dml(
        h,
        "INSERT INTO 在庫 VALUES ('A', 25) ON CONFLICT (品番) DO UPDATE SET 数量 = EXCLUDED.数量",
    );
    let n = dml(h, "UPDATE 在庫 SET 数量 = 数量 WHERE 数量 = 25");
    assert_eq!(n, 1, "UPSERT で更新されていない");
    conn::close(h).unwrap();
}

// ---------------------------------------------------------------------------
// インデックス
// ---------------------------------------------------------------------------

#[test]
fn index_variants_can_be_created_and_dropped() {
    let h = fresh_db("indexes.db");
    ddl(
        h,
        "CREATE TABLE 実績 (id INTEGER, コード VARCHAR, 金額 BIGINT, 日付 DATE)",
    );

    ddl(h, "CREATE INDEX idx_コード ON 実績 (コード)");
    ddl(h, "CREATE INDEX idx_複合 ON 実績 (コード, 日付)");
    ddl(h, "CREATE UNIQUE INDEX idx_一意 ON 実績 (id)");
    ddl(h, "CREATE INDEX IF NOT EXISTS idx_コード ON 実績 (コード)");
    ddl(h, "CREATE INDEX idx_式 ON 実績 (lower(コード))");
    ddl(h, "DROP INDEX idx_式");

    // UNIQUE インデックスが実際に効くこと
    dml(h, "INSERT INTO 実績 VALUES (1,'A',100,NULL)");
    assert!(!dml_err(h, "INSERT INTO 実績 VALUES (1,'B',200,NULL)").is_empty());

    conn::close(h).unwrap();
}

// ---------------------------------------------------------------------------
// ALTER TABLE — DuckDB の制約を含めて固定する
// ---------------------------------------------------------------------------

#[test]
fn alter_table_operations_work_without_indexes() {
    let h = fresh_db("alter_ok.db");
    ddl(
        h,
        "CREATE TABLE 素朴 (id INTEGER, 名前 VARCHAR, 区分 VARCHAR)",
    );

    ddl(h, "ALTER TABLE 素朴 ADD COLUMN 追加列 VARCHAR");
    ddl(
        h,
        "ALTER TABLE 素朴 ADD COLUMN 既定付 VARCHAR DEFAULT '未設定'",
    );
    ddl(h, "ALTER TABLE 素朴 DROP COLUMN 追加列");
    ddl(h, "ALTER TABLE 素朴 RENAME COLUMN 区分 TO 社員区分");
    ddl(
        h,
        "ALTER TABLE 素朴 ALTER COLUMN 社員区分 SET DEFAULT '正社員'",
    );
    ddl(h, "ALTER TABLE 素朴 ALTER COLUMN 社員区分 DROP DEFAULT");
    ddl(h, "ALTER TABLE 素朴 ALTER COLUMN 名前 SET NOT NULL");
    ddl(h, "ALTER TABLE 素朴 ALTER COLUMN 名前 DROP NOT NULL");
    ddl(h, "ALTER TABLE 素朴 ALTER COLUMN id TYPE BIGINT");
    ddl(h, "ALTER TABLE 素朴 RENAME TO 素朴2");

    conn::close(h).unwrap();
}

/// **DuckDB の制約**: インデックスが張られているテーブルは列の削除・改名ができない。
/// マニュアルに載せている回避策（インデックスを落として貼り直す）が有効なことも確認する。
#[test]
fn altering_a_table_with_an_index_needs_the_index_dropped_first() {
    let h = fresh_db("alter_index.db");
    ddl(
        h,
        "CREATE TABLE 索引付 (id INTEGER, 名前 VARCHAR, 不要列 VARCHAR)",
    );
    ddl(h, "CREATE INDEX idx_索引付 ON 索引付 (名前)");

    let err = ddl_err(h, "ALTER TABLE 索引付 DROP COLUMN 不要列");
    assert!(
        err.contains("depend"),
        "依存エラーになっていない（マニュアルの記述と食い違う）: {err}"
    );

    // マニュアルに書いた回避策
    ddl(h, "DROP INDEX idx_索引付");
    ddl(h, "ALTER TABLE 索引付 DROP COLUMN 不要列");
    ddl(h, "CREATE INDEX idx_索引付 ON 索引付 (名前)");

    conn::close(h).unwrap();
}

/// 主キーだけなら ALTER を妨げない（明示的なインデックスとは扱いが違う）。
#[test]
fn primary_key_alone_does_not_block_alter() {
    let h = fresh_db("alter_pk.db");
    ddl(
        h,
        "CREATE TABLE PK付 (id INTEGER PRIMARY KEY, 名前 VARCHAR, 不要列 VARCHAR)",
    );
    ddl(h, "ALTER TABLE PK付 DROP COLUMN 不要列");
    ddl(h, "ALTER TABLE PK付 ADD COLUMN 追加 VARCHAR");
    conn::close(h).unwrap();
}

/// **DuckDB の制約**: 制約の後付けはできない。テーブル作成時に書く必要がある。
#[test]
fn adding_a_constraint_after_creation_is_not_supported() {
    let h = fresh_db("add_constraint.db");
    ddl(h, "CREATE TABLE t (a INTEGER)");
    let err = ddl_err(h, "ALTER TABLE t ADD CONSTRAINT c CHECK (a > 0)");
    assert!(err.contains("support"), "未対応エラーになっていない: {err}");
    conn::close(h).unwrap();
}

// ---------------------------------------------------------------------------
// ビュー・シーケンス・スキーマ
// ---------------------------------------------------------------------------

#[test]
fn views_sequences_and_schemas_work() {
    let h = fresh_db("views.db");
    ddl(
        h,
        "CREATE TABLE 社員 (id INTEGER PRIMARY KEY, 氏名 VARCHAR, 在籍 BOOLEAN)",
    );
    dml(
        h,
        "INSERT INTO 社員 VALUES (1,'山田',true),(2,'鈴木',false)",
    );

    ddl(
        h,
        "CREATE VIEW 在籍者 AS SELECT id, 氏名 FROM 社員 WHERE 在籍",
    );
    ddl(
        h,
        "CREATE OR REPLACE VIEW 在籍者 AS SELECT id, 氏名 FROM 社員 WHERE 在籍 = true",
    );
    conn::with_conn(h, |s| query::query(s, "SELECT * FROM 在籍者"))
        .unwrap()
        .expect("ビューを SELECT できない")
        .clear();
    ddl(h, "DROP VIEW 在籍者");

    ddl(h, "CREATE SEQUENCE 連番 START 1");
    ddl(
        h,
        "CREATE TABLE 伝票 (id INTEGER DEFAULT nextval('連番'), 内容 VARCHAR)",
    );
    dml(h, "INSERT INTO 伝票 (内容) VALUES ('一件目'),('二件目')");

    ddl(h, "CREATE SCHEMA 業務");
    ddl(h, "CREATE TABLE 業務.受注 (id INTEGER PRIMARY KEY)");

    ddl(h, "COMMENT ON TABLE 社員 IS '社員マスタ'");

    conn::close(h).unwrap();
}

// ---------------------------------------------------------------------------
// 日付の「現在値」
// ---------------------------------------------------------------------------

/// CURRENT_DATE と today() は ICU 拡張を必要とし、外部アクセスを塞いでいる本 DLL では使えない。
/// マニュアルではこれを明記し、代替を案内している。その代替が本当に動くことを確認する。
#[test]
fn current_date_needs_icu_but_documented_alternatives_work() {
    let h = fresh_db("dates.db");

    // 使えないもの。エラーメッセージが icu を案内していること。
    for sql in ["SELECT CURRENT_DATE", "SELECT today()"] {
        let err = match conn::with_conn(h, |s| query::query(s, sql)).unwrap() {
            Ok(mut v) => {
                v.clear();
                panic!("{sql} が通った。マニュアルの記述を見直すこと");
            }
            Err(e) => e,
        };
        assert!(err.contains("icu"), "icu を案内していない: {err}");
    }

    // マニュアルで案内している代替
    for sql in [
        "SELECT now()",
        "SELECT CURRENT_TIMESTAMP",
        "SELECT get_current_timestamp()",
        "SELECT CAST(now() AS TIMESTAMP)",
        "SELECT CAST(CAST(now() AS TIMESTAMP) AS DATE)",
        "SELECT strftime(now(), '%Y-%m-%d')",
    ] {
        conn::with_conn(h, |s| query::query(s, sql))
            .unwrap()
            .unwrap_or_else(|e| panic!("代替として案内している {sql} が動かない: {e}"))
            .clear();
    }

    // 既定値としても使えること
    ddl(
        h,
        "CREATE TABLE 記録 (id INTEGER, 登録日 DATE DEFAULT CAST(CAST(now() AS TIMESTAMP) AS DATE))",
    );
    dml(h, "INSERT INTO 記録 (id) VALUES (1)");

    conn::close(h).unwrap();
}
