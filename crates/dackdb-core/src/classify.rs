//! SQL 文の種別判定と権限ゲート（防御の層2）。
//!
//! # なぜこれが必要か
//!
//! 層1（DuckDB エンジンの `access_mode=READ_ONLY`）は SQL からは突破不可能で、
//! 階層① 読み取り専用はそれだけでほぼ完結する。しかし DuckDB には
//! 「DML は可、DDL は不可」というモードが存在しないため、**階層② 読み書き可は
//! ここで実装するしかない**。
//!
//! # なぜ正規表現ではなく DuckDB のパーサを使うのか
//!
//! 先頭キーワードを見るだけの判定は下記のような入力で破綻する：
//!
//! - `WITH x AS (SELECT 1) INSERT INTO t SELECT * FROM x` — 先頭は WITH だが INSERT
//! - `SELECT 1; DROP TABLE t;` — 複数文の後半に DDL が紛れ込む
//! - `/* SELECT */ DROP TABLE t` — コメントによる偽装
//!
//! `duckdb_extract_statements` で DuckDB 自身に文を分割させ、
//! `duckdb_prepared_statement_type` で DuckDB 自身に種別を判定させれば、
//! パーサの解釈と実行の解釈が必ず一致する。

// bindgen が生成する定数名は `duckdb_statement_type_DUCKDB_STATEMENT_TYPE_*` で
// Rust の命名規則に反するが、こちらでは変えられない。
#![allow(non_upper_case_globals)]

use crate::level::Level;
use crate::raw::{Connection, Extracted};
use libduckdb_sys as ffi;

/// どの公開 API から呼ばれたか。API ごとの追加制限に使う。
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ApiKind {
    /// `DackQuery` / `DackQueryParams`：結果セットを返す読み取り。
    Query,
    /// `DackExecute` / `DackExecuteParams`：DML。
    Execute,
    /// `DackExecuteDDL`：スキーマ変更。管理者のみ。
    ExecuteDdl,
}

impl ApiKind {
    fn api_name(self) -> &'static str {
        match self {
            ApiKind::Query => "DackQuery",
            ApiKind::Execute => "DackExecute",
            ApiKind::ExecuteDdl => "DackExecuteDDL",
        }
    }
}

/// 文の種別を日本語名にする（エラーメッセージ用）。
pub fn statement_type_name(t: ffi::duckdb_statement_type) -> &'static str {
    use ffi::*;
    match t {
        duckdb_statement_type_DUCKDB_STATEMENT_TYPE_SELECT => "SELECT（参照）",
        duckdb_statement_type_DUCKDB_STATEMENT_TYPE_INSERT => "INSERT（挿入）",
        duckdb_statement_type_DUCKDB_STATEMENT_TYPE_UPDATE => "UPDATE（更新）",
        duckdb_statement_type_DUCKDB_STATEMENT_TYPE_DELETE => "DELETE（削除）",
        duckdb_statement_type_DUCKDB_STATEMENT_TYPE_MERGE_INTO => "MERGE INTO（併合）",
        duckdb_statement_type_DUCKDB_STATEMENT_TYPE_EXPLAIN => "EXPLAIN（実行計画）",
        duckdb_statement_type_DUCKDB_STATEMENT_TYPE_CREATE => "CREATE（作成）",
        duckdb_statement_type_DUCKDB_STATEMENT_TYPE_CREATE_FUNC => "CREATE FUNCTION（関数作成）",
        duckdb_statement_type_DUCKDB_STATEMENT_TYPE_DROP => "DROP（削除）",
        duckdb_statement_type_DUCKDB_STATEMENT_TYPE_ALTER => "ALTER（変更）",
        duckdb_statement_type_DUCKDB_STATEMENT_TYPE_ATTACH => "ATTACH（DB 接続）",
        duckdb_statement_type_DUCKDB_STATEMENT_TYPE_DETACH => "DETACH（DB 切断）",
        duckdb_statement_type_DUCKDB_STATEMENT_TYPE_COPY => "COPY（外部入出力）",
        duckdb_statement_type_DUCKDB_STATEMENT_TYPE_COPY_DATABASE => "COPY DATABASE",
        duckdb_statement_type_DUCKDB_STATEMENT_TYPE_EXPORT => "EXPORT（書き出し）",
        duckdb_statement_type_DUCKDB_STATEMENT_TYPE_PRAGMA => "PRAGMA（設定）",
        duckdb_statement_type_DUCKDB_STATEMENT_TYPE_SET => "SET（設定変更）",
        duckdb_statement_type_DUCKDB_STATEMENT_TYPE_VARIABLE_SET => "SET VARIABLE（変数設定）",
        duckdb_statement_type_DUCKDB_STATEMENT_TYPE_LOAD => "LOAD / INSTALL（拡張）",
        duckdb_statement_type_DUCKDB_STATEMENT_TYPE_TRANSACTION => "トランザクション制御",
        duckdb_statement_type_DUCKDB_STATEMENT_TYPE_PREPARE => "PREPARE（文の準備）",
        duckdb_statement_type_DUCKDB_STATEMENT_TYPE_EXECUTE => "EXECUTE（準備文の実行）",
        duckdb_statement_type_DUCKDB_STATEMENT_TYPE_VACUUM => "VACUUM（再構成）",
        duckdb_statement_type_DUCKDB_STATEMENT_TYPE_ANALYZE => "ANALYZE（統計収集）",
        duckdb_statement_type_DUCKDB_STATEMENT_TYPE_CALL => "CALL（関数呼び出し）",
        duckdb_statement_type_DUCKDB_STATEMENT_TYPE_MULTI => "複合文",
        duckdb_statement_type_DUCKDB_STATEMENT_TYPE_UPDATE_EXTENSIONS => "UPDATE EXTENSIONS",
        duckdb_statement_type_DUCKDB_STATEMENT_TYPE_RELATION => "RELATION",
        duckdb_statement_type_DUCKDB_STATEMENT_TYPE_EXTENSION => "EXTENSION",
        duckdb_statement_type_DUCKDB_STATEMENT_TYPE_LOGICAL_PLAN => "LOGICAL PLAN",
        _ => "不明な種別",
    }
}

/// 読み取り系（どの階層でも許可される）。
fn is_read_only_type(t: ffi::duckdb_statement_type) -> bool {
    use ffi::*;
    matches!(
        t,
        duckdb_statement_type_DUCKDB_STATEMENT_TYPE_SELECT
            | duckdb_statement_type_DUCKDB_STATEMENT_TYPE_EXPLAIN
    )
}

/// DML（行の書き換え。スキーマは変えない）。
fn is_dml_type(t: ffi::duckdb_statement_type) -> bool {
    use ffi::*;
    matches!(
        t,
        duckdb_statement_type_DUCKDB_STATEMENT_TYPE_INSERT
            | duckdb_statement_type_DUCKDB_STATEMENT_TYPE_UPDATE
            | duckdb_statement_type_DUCKDB_STATEMENT_TYPE_DELETE
            | duckdb_statement_type_DUCKDB_STATEMENT_TYPE_MERGE_INTO
            | duckdb_statement_type_DUCKDB_STATEMENT_TYPE_TRANSACTION
    )
}

/// この権限レベルでこの文種別を実行してよいか。
///
/// **明示的な許可リスト方式**（拒否リストではない）。DuckDB に新しい文種別が
/// 追加されたとき、拒否リスト方式だと自動的に許可されてしまい権限が漏れる。
/// 許可リストなら未知の種別は既定で拒否される（フェイルクローズ）。
///
/// なお `PREPARE` / `EXECUTE` は階層② でも許可しない。`PREPARE x AS <DDL>` の形で
/// DDL を後から実行できてしまう余地を残さないため。パラメータ付き実行が必要な場合は
/// `DackQueryParams` / `DackExecuteParams`（`?` バインド）を使うこと。
pub fn is_allowed(level: Level, t: ffi::duckdb_statement_type) -> bool {
    match level {
        Level::Read => is_read_only_type(t),
        Level::ReadWrite => is_read_only_type(t) || is_dml_type(t),
        Level::Admin => t != ffi::duckdb_statement_type_DUCKDB_STATEMENT_TYPE_INVALID,
    }
}

/// API ごとの追加制限。
///
/// `DackQuery` は全階層で SELECT / EXPLAIN のみに絞る。書き込みは必ず
/// `DackExecute` / `DackExecuteDDL` を経由させることで、**VBA のコードを読んだだけで
/// 副作用の有無が分かる**ようにするため。
fn api_allows(api: ApiKind, t: ffi::duckdb_statement_type) -> bool {
    match api {
        ApiKind::Query => is_read_only_type(t),
        ApiKind::Execute => is_read_only_type(t) || is_dml_type(t),
        ApiKind::ExecuteDdl => true,
    }
}

/// SQL 全体を検査する。**1 文でも許可されない文が含まれていたら全体を拒否**する。
///
/// 部分実行を許すと `INSERT ...; DROP TABLE t;` で INSERT だけ通ってから
/// 失敗する、といった中途半端な状態が起きるため、実行前に全文を検査する。
pub fn check(
    level: Level,
    api: ApiKind,
    conn: &Connection,
    sql: &str,
) -> Result<(), String> {
    let extracted = Extracted::extract(conn, sql)?;
    let n = extracted.count();

    for i in 0..n {
        // prepare に失敗した場合は種別を確定できないので拒否する（フェイルクローズ）。
        let prepared = extracted.prepare(conn, i)?;
        check_statement_type(level, api, prepared.statement_type())?;
    }
    Ok(())
}

/// 判定済みの文種別に対する権限チェック。
///
/// [`check`] のほか、パラメータ付き実行のように prepared statement を
/// 自前で作る経路からも使う。**権限判定はここ 1 箇所に集約する**。
pub fn check_statement_type(
    level: Level,
    api: ApiKind,
    t: ffi::duckdb_statement_type,
) -> Result<(), String> {
    let type_name = statement_type_name(t);

    if !is_allowed(level, t) {
        return Err(format!(
            "この DLL は{}（{}）です。{} は実行できません。{}",
            level.description_ja(),
            level.name(),
            type_name,
            upgrade_hint(level, t)
        ));
    }
    if !api_allows(api, t) {
        return Err(format!(
            "{} では {} を実行できません。{}",
            api.api_name(),
            type_name,
            api_hint(t)
        ));
    }
    Ok(())
}

/// 上位の DLL を使えば実行できることを案内する。
fn upgrade_hint(level: Level, t: ffi::duckdb_statement_type) -> &'static str {
    match level {
        Level::Read => {
            if is_dml_type(t) {
                "書き込みには dackdb_rw.dll（読み書き可）を使ってください。"
            } else {
                "スキーマ変更には dackdb_admin.dll（管理者）を使ってください。"
            }
        }
        Level::ReadWrite => "スキーマ変更には dackdb_admin.dll（管理者）を使ってください。",
        Level::Admin => "",
    }
}

fn api_hint(t: ffi::duckdb_statement_type) -> &'static str {
    if is_dml_type(t) {
        "DackExecute を使ってください。"
    } else {
        "DackExecuteDDL を使ってください。"
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use ffi::*;

    const SELECT: duckdb_statement_type = duckdb_statement_type_DUCKDB_STATEMENT_TYPE_SELECT;
    const INSERT: duckdb_statement_type = duckdb_statement_type_DUCKDB_STATEMENT_TYPE_INSERT;
    const CREATE: duckdb_statement_type = duckdb_statement_type_DUCKDB_STATEMENT_TYPE_CREATE;
    const DROP: duckdb_statement_type = duckdb_statement_type_DUCKDB_STATEMENT_TYPE_DROP;
    const ATTACH: duckdb_statement_type = duckdb_statement_type_DUCKDB_STATEMENT_TYPE_ATTACH;
    const SET: duckdb_statement_type = duckdb_statement_type_DUCKDB_STATEMENT_TYPE_SET;
    const PRAGMA: duckdb_statement_type = duckdb_statement_type_DUCKDB_STATEMENT_TYPE_PRAGMA;
    const LOAD: duckdb_statement_type = duckdb_statement_type_DUCKDB_STATEMENT_TYPE_LOAD;
    const COPY: duckdb_statement_type = duckdb_statement_type_DUCKDB_STATEMENT_TYPE_COPY;
    const PREPARE: duckdb_statement_type = duckdb_statement_type_DUCKDB_STATEMENT_TYPE_PREPARE;
    const EXECUTE: duckdb_statement_type = duckdb_statement_type_DUCKDB_STATEMENT_TYPE_EXECUTE;
    const MULTI: duckdb_statement_type = duckdb_statement_type_DUCKDB_STATEMENT_TYPE_MULTI;
    const INVALID: duckdb_statement_type = duckdb_statement_type_DUCKDB_STATEMENT_TYPE_INVALID;

    #[test]
    fn read_level_allows_only_select_and_explain() {
        assert!(is_allowed(Level::Read, SELECT));
        for t in [INSERT, CREATE, DROP, ATTACH, SET, PRAGMA, LOAD, COPY, MULTI] {
            assert!(!is_allowed(Level::Read, t), "Read が {t} を許可している");
        }
    }

    #[test]
    fn readwrite_allows_dml_but_never_ddl() {
        assert!(is_allowed(Level::ReadWrite, SELECT));
        assert!(is_allowed(Level::ReadWrite, INSERT));
        for t in [CREATE, DROP, ATTACH, SET, PRAGMA, LOAD, COPY, MULTI] {
            assert!(!is_allowed(Level::ReadWrite, t), "ReadWrite が {t} を許可している");
        }
    }

    /// PREPARE / EXECUTE を階層② で通すと `PREPARE x AS DROP TABLE t` の余地が残る。
    #[test]
    fn prepare_and_execute_are_denied_below_admin() {
        for level in [Level::Read, Level::ReadWrite] {
            assert!(!is_allowed(level, PREPARE), "{level:?} が PREPARE を許可している");
            assert!(!is_allowed(level, EXECUTE), "{level:?} が EXECUTE を許可している");
        }
    }

    #[test]
    fn admin_allows_everything_except_invalid() {
        for t in [SELECT, INSERT, CREATE, DROP, ATTACH, SET, PRAGMA, LOAD, COPY, MULTI] {
            assert!(is_allowed(Level::Admin, t), "Admin が {t} を拒否している");
        }
        assert!(!is_allowed(Level::Admin, INVALID));
    }

    /// 許可リスト方式なので、未知の（将来追加される）文種別は既定で拒否される。
    #[test]
    fn unknown_future_statement_types_fail_closed() {
        let future: duckdb_statement_type = 9999;
        assert!(!is_allowed(Level::Read, future));
        assert!(!is_allowed(Level::ReadWrite, future));
    }

    #[test]
    fn dack_query_never_permits_writes_even_for_admin() {
        assert!(api_allows(ApiKind::Query, SELECT));
        assert!(!api_allows(ApiKind::Query, INSERT));
        assert!(!api_allows(ApiKind::Query, CREATE));
        // DackExecute は DML まで
        assert!(api_allows(ApiKind::Execute, INSERT));
        assert!(!api_allows(ApiKind::Execute, CREATE));
        // DackExecuteDDL は何でも（レベル判定で別途絞られる）
        assert!(api_allows(ApiKind::ExecuteDdl, CREATE));
    }

    #[test]
    fn every_statement_type_has_a_japanese_name() {
        for t in 1..=30u32 {
            assert_ne!(
                statement_type_name(t),
                "不明な種別",
                "種別 {t} に日本語名が無い"
            );
        }
    }
}
