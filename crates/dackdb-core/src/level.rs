//! 権限レベルの定義。
//!
//! 重要：レベルは Cargo の feature ではなく、この enum の定数を各 cdylib クレートから
//! `export_dackdb_ffi!` マクロに渡すことで決める。feature にするとワークスペース一括
//! ビルド時に feature 統合が起きて 3 つの DLL がすべて管理者権限になってしまう。

/// DLL が持つ権限レベル。
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord)]
pub enum Level {
    /// ① 読み取り専用。SELECT のみ。
    Read,
    /// ② 読み書き可。DML は可、DDL は不可。
    ReadWrite,
    /// ③ 管理者。DB 作成・DDL・スキーマ出力まですべて。
    Admin,
}

impl Level {
    /// VBA へ返す表示名。
    pub const fn name(self) -> &'static str {
        match self {
            Level::Read => "READ_ONLY",
            Level::ReadWrite => "READ_WRITE",
            Level::Admin => "ADMIN",
        }
    }

    /// 日本語の説明。エラーメッセージに使う。
    pub const fn description_ja(self) -> &'static str {
        match self {
            Level::Read => "読み取り専用",
            Level::ReadWrite => "読み書き可（DDL 不可）",
            Level::Admin => "管理者（全機能）",
        }
    }

    /// 書き込み系 API（DackExecute / DackAppendArray / トランザクション）が使えるか。
    pub const fn allows_write(self) -> bool {
        matches!(self, Level::ReadWrite | Level::Admin)
    }

    /// DDL 系 API（DackCreateDatabase / DackExecuteDDL / DackExportSchema）が使えるか。
    pub const fn allows_ddl(self) -> bool {
        matches!(self, Level::Admin)
    }
}

/// このレベルで許可されるエクスポート関数の一覧。`DackCapabilities` が返す。
pub fn capabilities(level: Level) -> Vec<&'static str> {
    let mut v = vec![
        "DackVersion",
        "DackCapabilities",
        "DackOpen",
        "DackClose",
        "DackQuery",
        "DackQueryParams",
        "DackListTables",
        "DackDescribe",
    ];
    if level.allows_write() {
        v.extend_from_slice(&[
            "DackExecute",
            "DackExecuteParams",
            "DackAppendArray",
            "DackBegin",
            "DackCommit",
            "DackRollback",
        ]);
    }
    if level.allows_ddl() {
        v.extend_from_slice(&[
            "DackCreateDatabase",
            "DackExecuteDDL",
            "DackExportSchema",
            "DackCheckpoint",
        ]);
    }
    v
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn read_level_has_no_write_or_ddl() {
        assert!(!Level::Read.allows_write());
        assert!(!Level::Read.allows_ddl());
    }

    #[test]
    fn readwrite_allows_write_but_not_ddl() {
        assert!(Level::ReadWrite.allows_write());
        assert!(!Level::ReadWrite.allows_ddl());
    }

    #[test]
    fn admin_allows_everything() {
        assert!(Level::Admin.allows_write());
        assert!(Level::Admin.allows_ddl());
    }

    #[test]
    fn capabilities_grow_monotonically() {
        let r = capabilities(Level::Read);
        let rw = capabilities(Level::ReadWrite);
        let a = capabilities(Level::Admin);
        assert!(r.len() < rw.len() && rw.len() < a.len());
        // 下位レベルの機能は上位レベルにすべて含まれる（VBA コードの上位互換性）
        assert!(r.iter().all(|f| rw.contains(f)));
        assert!(rw.iter().all(|f| a.contains(f)));
    }
}
