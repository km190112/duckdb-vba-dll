//! 公開 API の実装本体。FFI マクロはここを呼ぶだけの薄い殻。
//!
//! すべての関数は `Result<VARIANT, String>` を返す。成功時の VARIANT が
//! VBA の `ByRef result As Variant` に書き込まれ、失敗時は文字列（エラーメッセージ）が
//! 同じ `result` に書き込まれる。VBA 側から見ると常に同じ形になる。

use crate::conn::{self, ConnState, OpenOptions};
use crate::level::{self, Level};
use crate::oleaut::VARIANT;
use crate::query;

// ---------------------------------------------------------------------------
// 戻り値コード
// ---------------------------------------------------------------------------

/// 成功。
pub const DACK_OK: i32 = 0;
/// 一般的なエラー。詳細は `result` の文字列に入る。
pub const DACK_E_GENERAL: i32 = -1;
/// この権限レベルでは使えない関数を呼んだ。
pub const DACK_E_FORBIDDEN: i32 = -403;
/// Rust 側で panic が起きた（本来起きてはいけない。バグ報告対象）。
pub const DACK_E_PANIC: i32 = -999;

/// 何も返さない成功（`DackClose` など）を表す VARIANT。
fn ok_empty() -> VARIANT {
    VARIANT::empty()
}

/// その権限レベルでは使えない関数、というエラー。
fn forbidden(level: Level, func: &str, needed: &str) -> String {
    format!(
        "{func} はこの DLL では使えません。この DLL は{}（{}）です。\n\
         {needed}",
        level.description_ja(),
        level.name()
    )
}

// ---------------------------------------------------------------------------
// 全階層で使える関数
// ---------------------------------------------------------------------------

pub fn version(level: Level) -> Result<VARIANT, String> {
    Ok(VARIANT::bstr(&crate::version_string(level)))
}

pub fn capabilities(level: Level) -> Result<VARIANT, String> {
    Ok(VARIANT::bstr(&level::capabilities(level).join(", ")))
}

/// データベースを開く。成功時の `result` は接続ハンドル（`LongLong`）。
pub fn open(level: Level, path: &str) -> Result<VARIANT, String> {
    let h = conn::open(level, path, OpenOptions::default())?;
    Ok(VARIANT::i64(h))
}

pub fn close(handle: i64) -> Result<VARIANT, String> {
    conn::close(handle)?;
    Ok(ok_empty())
}

pub fn query(handle: i64, sql: &str) -> Result<VARIANT, String> {
    with(handle, |s| query::query(s, sql))
}

/// パラメータ付き SELECT。`?` に値をバインドするので SQL インジェクションが起きない。
///
/// # Safety
/// `params` は VBA が所有する有効な VARIANT（配列）へのポインタであること。
pub unsafe fn query_params(
    handle: i64,
    sql: &str,
    params: *const VARIANT,
) -> Result<VARIANT, String> {
    if params.is_null() {
        return Err(
            "パラメータが空です。Array(...) か Range(...).Value を渡してください。".to_string(),
        );
    }
    let grid = crate::inbound::read_input_grid(&*params, "パラメータ")?;
    with(handle, |s| query::query_params(s, sql, &grid))
}

pub fn list_tables(handle: i64) -> Result<VARIANT, String> {
    with(handle, query::list_tables)
}

pub fn describe(handle: i64, table: &str) -> Result<VARIANT, String> {
    with(handle, |s| query::describe(s, table))
}

// ---------------------------------------------------------------------------
// 階層② 以上
// ---------------------------------------------------------------------------

pub fn execute(level: Level, handle: i64, sql: &str) -> Result<VARIANT, String> {
    if !level.allows_write() {
        return Err(forbidden(
            level,
            "DackExecute",
            "書き込みには dackdb_rw.dll（読み書き可）を使ってください。",
        ));
    }
    with(handle, |s| query::execute(s, sql).map(VARIANT::i64))
}

/// パラメータ付き DML。
///
/// # Safety
/// `params` は VBA が所有する有効な VARIANT（配列）へのポインタであること。
pub unsafe fn execute_params(
    level: Level,
    handle: i64,
    sql: &str,
    params: *const VARIANT,
) -> Result<VARIANT, String> {
    if !level.allows_write() {
        return Err(forbidden(
            level,
            "DackExecuteParams",
            "書き込みには dackdb_rw.dll（読み書き可）を使ってください。",
        ));
    }
    if params.is_null() {
        return Err(
            "パラメータが空です。Array(...) か Range(...).Value を渡してください。".to_string(),
        );
    }
    let grid = crate::inbound::read_input_grid(&*params, "パラメータ")?;
    with(handle, |s| {
        query::execute_params(s, sql, &grid).map(VARIANT::i64)
    })
}

/// シート範囲をテーブルへ一括投入する。INSERT 文の反復より 1〜2 桁速い。
///
/// # Safety
/// `data` は VBA が所有する有効な VARIANT（2 次元配列）へのポインタであること。
pub unsafe fn append_array(
    level: Level,
    handle: i64,
    table: &str,
    data: *const VARIANT,
) -> Result<VARIANT, String> {
    if !level.allows_write() {
        return Err(forbidden(
            level,
            "DackAppendArray",
            "書き込みには dackdb_rw.dll（読み書き可）を使ってください。",
        ));
    }
    if data.is_null() {
        return Err(
            "データが空です。Range(\"A2:F100\").Value のように範囲を渡してください。".to_string(),
        );
    }
    let grid = crate::inbound::read_input_grid(&*data, "データ")?;
    with(handle, |s| {
        crate::append::append_array(s, table, &grid).map(VARIANT::i64)
    })
}

pub fn begin(level: Level, handle: i64) -> Result<VARIANT, String> {
    transaction_stmt(level, handle, "BEGIN TRANSACTION", "DackBegin")
}

pub fn commit(level: Level, handle: i64) -> Result<VARIANT, String> {
    transaction_stmt(level, handle, "COMMIT", "DackCommit")
}

pub fn rollback(level: Level, handle: i64) -> Result<VARIANT, String> {
    transaction_stmt(level, handle, "ROLLBACK", "DackRollback")
}

fn transaction_stmt(level: Level, handle: i64, sql: &str, func: &str) -> Result<VARIANT, String> {
    if !level.allows_write() {
        return Err(forbidden(
            level,
            func,
            "トランザクションには dackdb_rw.dll（読み書き可）以上を使ってください。",
        ));
    }
    with(handle, |s| {
        s.conn.query(sql)?;
        Ok(ok_empty())
    })
}

// ---------------------------------------------------------------------------
// 階層③ のみ
// ---------------------------------------------------------------------------

/// 新しい `dack.db` を作成する。既存ファイルがあればエラーにする
/// （黙って上書きすると業務データが消えるため）。
pub fn create_database(level: Level, path: &str) -> Result<VARIANT, String> {
    if !level.allows_ddl() {
        return Err(forbidden(
            level,
            "DackCreateDatabase",
            "データベースの作成には dackdb_admin.dll（管理者）を使ってください。",
        ));
    }
    let path = path.trim();
    if path.is_empty() {
        return Err("作成するデータベースのパスが空です。".to_string());
    }
    if std::path::Path::new(path).exists() {
        return Err(format!(
            "ファイルが既に存在します: {path}\n\
             既存のデータベースを開くには DackOpen を使ってください。\
             上書きが必要な場合は先にファイルを削除してください。"
        ));
    }
    let h = conn::open(level, path, OpenOptions::default())?;
    Ok(VARIANT::i64(h))
}

pub fn execute_ddl(level: Level, handle: i64, sql: &str) -> Result<VARIANT, String> {
    if !level.allows_ddl() {
        return Err(forbidden(
            level,
            "DackExecuteDDL",
            "スキーマ変更には dackdb_admin.dll（管理者）を使ってください。",
        ));
    }
    with(handle, |s| query::execute_ddl(s, sql).map(VARIANT::i64))
}

pub fn export_schema(level: Level, handle: i64, format: &str) -> Result<VARIANT, String> {
    if !level.allows_ddl() {
        return Err(forbidden(
            level,
            "DackExportSchema",
            "スキーマ情報の出力には dackdb_admin.dll（管理者）を使ってください。",
        ));
    }
    with(handle, |s| crate::schema::export(s, format))
}

pub fn checkpoint(level: Level, handle: i64) -> Result<VARIANT, String> {
    if !level.allows_ddl() {
        return Err(forbidden(
            level,
            "DackCheckpoint",
            "CHECKPOINT には dackdb_admin.dll（管理者）を使ってください。",
        ));
    }
    with(handle, |s| {
        s.conn.query("CHECKPOINT")?;
        Ok(ok_empty())
    })
}

// ---------------------------------------------------------------------------
// 共通ヘルパ
// ---------------------------------------------------------------------------

/// ハンドルから接続を引いて処理を行う。`with_conn` の二重 Result を畳む。
fn with(
    handle: i64,
    f: impl FnOnce(&ConnState) -> Result<VARIANT, String>,
) -> Result<VARIANT, String> {
    conn::with_conn(handle, f)?
}
