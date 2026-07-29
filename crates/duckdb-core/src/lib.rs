//! Excel VBA から DuckDB を操作するための権限階層付きコアロジック。
//!
//! # 全体像
//!
//! `dackdb-r` / `dackdb-rw` / `dackdb-admin` の 3 つの cdylib クレートは、
//! いずれもこのクレートの [`export_dackdb_ffi!`] マクロに [`Level`] 定数を
//! 渡すだけの薄い殻である。ロジックはすべてここにある。
//!
//! # 権限は 2 層で強制する
//!
//! - **層1**：接続時の DuckDB 設定（`access_mode=READ_ONLY` など）。SQL からは突破できない。
//!   → [`conn`] モジュール
//! - **層2**：DuckDB のパーサによる文種別の許可リスト判定。
//!   → [`classify`] モジュール
//!
//! # VBA との境界
//!
//! 入力文字列は `StrPtr()` の生 UTF-16 ポインタで受け取り、出力はすべて
//! `ByRef ... As Variant` に書き込む。VBA の `Declare` による ANSI マーシャリングを
//! 一切経由しないため、日本語が文字化けしない。
//! → [`oleaut`] / [`variant`] モジュール

pub mod api;
pub mod append;
pub mod classify;
pub mod conn;
pub mod ffi;
pub mod inbound;
pub mod level;
pub mod oleaut;
pub mod query;
pub mod raw;
pub mod schema;
pub mod value;
pub mod variant;

pub use level::Level;

/// このクレートの版数。
pub const VERSION: &str = env!("CARGO_PKG_VERSION");

/// DuckDB のライブラリ版数を返す。
pub fn duckdb_version() -> String {
    unsafe {
        let p = libduckdb_sys::duckdb_library_version();
        if p.is_null() {
            return "unknown".to_string();
        }
        std::ffi::CStr::from_ptr(p).to_string_lossy().into_owned()
    }
}

/// `DackVersion` が返す文字列。
pub fn version_string(level: Level) -> String {
    format!(
        "dackdb {} / DuckDB {} / 権限: {} ({})",
        VERSION,
        duckdb_version(),
        level.name(),
        level.description_ja()
    )
}
