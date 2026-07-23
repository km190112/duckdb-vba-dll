//! 接続の生成（防御の層1）とハンドルレジストリ。
//!
//! # 層1：DuckDB エンジンレベルの権限
//!
//! 接続時の `duckdb_config` で権限を決める。ここで設定した `access_mode=READ_ONLY` は
//! **SQL からは突破できない**。DuckDB 自身が
//! `Cannot execute statement of type 'CREATE' on database which is attached in read-only mode!`
//! を返して全書き込み・DDL・書き込み ATTACH を拒否する。
//!
//! `lock_configuration=true` を**必ず最後に**設定するのが要点。これを先に設定すると
//! 以降の `duckdb_set_config` が効かなくなり、逆に設定しないと利用者が
//! `SET access_mode='READ_WRITE'` で昇格できてしまう。
//!
//! # ハンドルは生ポインタを VBA に渡さない
//!
//! VBA から壊れた `LongLong` が渡ってきても Excel をクラッシュさせないため、
//! 連番の整数ハンドルとレジストリの間接参照を挟む。

use crate::level::Level;
use crate::raw::{Config, Connection, Database};
use std::collections::HashMap;
use std::sync::atomic::{AtomicI64, Ordering};
use std::sync::{Mutex, OnceLock};

/// `DackOpen` の追加オプション。管理者 DLL のみ意味を持つ。
#[derive(Debug, Clone, Copy, Default)]
pub struct OpenOptions {
    /// 外部ファイルアクセス（CSV/Parquet 読み書き、`COPY`、拡張の導入）を許可するか。
    /// 階層① ② では常に無効。
    pub allow_external_access: bool,
}

/// 1 つの接続の状態。
pub struct ConnState {
    /// `Database` は `Connection` より後に破棄される必要がある。
    /// Rust の構造体フィールドは宣言順に drop されるため、conn を先に宣言する。
    pub conn: Connection,
    #[allow(dead_code)]
    db: Database,
    pub level: Level,
    pub path: String,
}

fn registry() -> &'static Mutex<HashMap<i64, ConnState>> {
    static R: OnceLock<Mutex<HashMap<i64, ConnState>>> = OnceLock::new();
    R.get_or_init(|| Mutex::new(HashMap::new()))
}

/// レジストリのロックを取る。**poisoning から必ず復帰する**。
///
/// FFI 境界では panic を `catch_unwind` で捕まえて VBA にエラーを返すが、
/// その panic がロック保持中に起きると Mutex が poisoned になる。素直に
/// `lock()?` にしていると、以降**すべての** DLL 呼び出しが失敗し続け、
/// Excel を再起動するまで復旧しない。
///
/// ここで保護しているのは「ハンドル → 接続」の HashMap だけであり、
/// クエリ実行中の panic でこのマップ自体が壊れることはない。よって
/// poisoning を無視して中身を取り出すのが正しい。
fn lock_registry() -> std::sync::MutexGuard<'static, HashMap<i64, ConnState>> {
    registry()
        .lock()
        .unwrap_or_else(|poisoned| poisoned.into_inner())
}

/// 0 は「無効なハンドル」を表すので 1 から始める。
static NEXT_HANDLE: AtomicI64 = AtomicI64::new(1);

/// 権限レベルに対応する DuckDB 設定を組み立てる。
///
/// **明示的にすべての関連設定を書き下している**。既定値に頼ると、
/// DuckDB の既定が将来変わったときに権限が静かに緩む可能性があるため。
fn build_config(level: Level, opts: OpenOptions) -> Result<Config, String> {
    let mut c = Config::new()?;

    match level {
        Level::Read => {
            c.set("access_mode", "READ_ONLY")?;
            c.set("enable_external_access", "false")?;
        }
        Level::ReadWrite => {
            c.set("access_mode", "READ_WRITE")?;
            c.set("enable_external_access", "false")?;
        }
        Level::Admin => {
            c.set("access_mode", "READ_WRITE")?;
            c.set(
                "enable_external_access",
                if opts.allow_external_access {
                    "true"
                } else {
                    "false"
                },
            )?;
        }
    }

    // 拡張の読み込みは全階層で塞ぐ。署名なし拡張を読めると任意コード実行になり、
    // 権限階層が意味を失うため、管理者 DLL でも許可しない。
    c.set("allow_unsigned_extensions", "false")?;

    // 階層① ② では拡張の自動導入・自動読み込みも無効化する。
    if level != Level::Admin {
        c.set("autoinstall_known_extensions", "false")?;
        c.set("autoload_known_extensions", "false")?;
    }

    // ---- 以降に設定を追加しないこと ----
    // lock_configuration は最後。これ以降の set_config は無視される。
    if level != Level::Admin {
        c.set("lock_configuration", "true")?;
    }

    Ok(c)
}

/// データベースを開き、ハンドルを返す。
pub fn open(level: Level, path: &str, opts: OpenOptions) -> Result<i64, String> {
    let path = path.trim();
    if path.is_empty() {
        return Err(
            "データベースのパスが空です。dack.db のフルパスを指定してください。".to_string(),
        );
    }

    // 読み取り専用でファイルが無い場合、DuckDB のエラーは分かりにくいので先に案内する。
    if level == Level::Read && path != ":memory:" && !std::path::Path::new(path).exists() {
        return Err(format!(
            "データベースファイルが見つかりません: {path}\n\
             読み取り専用 DLL は新しいファイルを作成できません。\
             作成するには dackdb_admin.dll の DackCreateDatabase を使ってください。"
        ));
    }

    let config = build_config(level, opts)?;
    let db = Database::open(path, &config)?;
    let conn = db.connect()?;

    let handle = NEXT_HANDLE.fetch_add(1, Ordering::SeqCst);
    let state = ConnState {
        conn,
        db,
        level,
        path: path.to_string(),
    };

    lock_registry().insert(handle, state);

    Ok(handle)
}

/// ハンドルを閉じる。既に閉じているハンドルは明示的なエラーにする
/// （二重 Close を黙って成功にすると VBA 側のバグが隠れるため）。
pub fn close(handle: i64) -> Result<(), String> {
    match lock_registry().remove(&handle) {
        Some(_) => Ok(()),
        None => Err(bad_handle_message(handle)),
    }
}

/// ハンドルに対応する接続を借りて処理を行う。
///
/// レジストリのロックを保持したまま `f` を呼ぶので、同一接続への同時アクセスは
/// 自動的に直列化される（VBA は基本的に単一スレッドだが、防御的に）。
pub fn with_conn<R>(handle: i64, f: impl FnOnce(&ConnState) -> R) -> Result<R, String> {
    let guard = lock_registry();
    match guard.get(&handle) {
        Some(state) => Ok(f(state)),
        None => Err(bad_handle_message(handle)),
    }
}

fn bad_handle_message(handle: i64) -> String {
    format!(
        "接続ハンドル {handle} は無効です。\
         DackOpen が成功していないか、既に DackClose 済みの可能性があります。"
    )
}

/// 開いている接続の数（テストと診断用）。
pub fn open_count() -> usize {
    lock_registry().len()
}

#[cfg(test)]
mod tests {
    use super::*;

    fn tmp_db(name: &str) -> std::path::PathBuf {
        let dir = std::env::temp_dir().join("dackdb-tests");
        std::fs::create_dir_all(&dir).unwrap();
        let p = dir.join(name);
        let _ = std::fs::remove_file(&p);
        p
    }

    /// 管理者権限で DB を作り、テーブルとデータを入れておく。
    fn seed(path: &str) -> i64 {
        let h = open(Level::Admin, path, OpenOptions::default()).expect("管理者で作成できるはず");
        with_conn(h, |s| {
            s.conn
                .query("CREATE TABLE 売上 (id INTEGER, 部署 VARCHAR, 金額 BIGINT)")
                .unwrap();
            s.conn
                .query("INSERT INTO 売上 VALUES (1, '営業部', 1000), (2, '技術部', 2000)")
                .unwrap();
        })
        .unwrap();
        h
    }

    #[test]
    fn read_only_engine_rejects_writes_regardless_of_sql_layer() {
        let path = tmp_db("readonly.db");
        let p = path.to_str().unwrap();
        let admin = seed(p);
        close(admin).unwrap();

        let h = open(Level::Read, p, OpenOptions::default()).unwrap();
        with_conn(h, |s| {
            // 層1（エンジン）だけで書き込みが拒否されることを確認する。
            // ここでは classify（層2）を通していないことが重要。
            let err = s
                .conn
                .query("INSERT INTO 売上 VALUES (3, '経理部', 3000)")
                .unwrap_err();
            assert!(
                err.contains("read-only") || err.contains("read only"),
                "エンジンが READ_ONLY を強制していない: {err}"
            );

            let err = s.conn.query("CREATE TABLE t2 (a INTEGER)").unwrap_err();
            assert!(
                err.contains("read-only") || err.contains("read only"),
                "{err}"
            );

            // 参照はできる
            s.conn.query("SELECT * FROM 売上").unwrap();
        })
        .unwrap();
        close(h).unwrap();
    }

    /// `lock_configuration=true` により `SET access_mode` での昇格ができないこと。
    /// これが破れると階層① ② の層1 防御が丸ごと無効になる。
    #[test]
    fn configuration_is_locked_against_privilege_escalation() {
        let path = tmp_db("locked.db");
        let p = path.to_str().unwrap();
        let admin = seed(p);
        close(admin).unwrap();

        let h = open(Level::Read, p, OpenOptions::default()).unwrap();
        with_conn(h, |s| {
            // 昇格の試み。成功してはいけない。
            let escalated = s.conn.query("SET access_mode='READ_WRITE'").is_ok();
            let still_readonly = s.conn.query("INSERT INTO 売上 VALUES (9, 'x', 1)").is_err();
            assert!(
                !escalated || still_readonly,
                "SET access_mode で読み書きに昇格できてしまった"
            );
        })
        .unwrap();
        close(h).unwrap();
    }

    #[test]
    fn readwrite_can_insert_but_engine_still_present() {
        let path = tmp_db("rw.db");
        let p = path.to_str().unwrap();
        let admin = seed(p);
        close(admin).unwrap();

        let h = open(Level::ReadWrite, p, OpenOptions::default()).unwrap();
        with_conn(h, |s| {
            s.conn
                .query("INSERT INTO 売上 VALUES (3, '経理部', 3000)")
                .unwrap();
        })
        .unwrap();
        close(h).unwrap();
    }

    #[test]
    fn read_level_reports_missing_file_clearly() {
        let err = open(
            Level::Read,
            "C:/存在しないフォルダ/無い.db",
            OpenOptions::default(),
        )
        .unwrap_err();
        assert!(err.contains("見つかりません"), "{err}");
        assert!(
            err.contains("dackdb_admin.dll"),
            "上位 DLL への案内が無い: {err}"
        );
    }

    #[test]
    fn empty_path_is_rejected() {
        let err = open(Level::Read, "   ", OpenOptions::default()).unwrap_err();
        assert!(err.contains("パスが空"), "{err}");
    }

    #[test]
    fn bad_handles_error_instead_of_crashing() {
        for h in [0i64, -1, i64::MAX, i64::MIN, 999_999] {
            let r = with_conn(h, |_| ());
            assert!(r.is_err(), "ハンドル {h} がエラーにならなかった");
            assert!(
                close(h).is_err(),
                "ハンドル {h} の close がエラーにならなかった"
            );
        }
    }

    #[test]
    fn double_close_is_an_error() {
        let path = tmp_db("double.db");
        let p = path.to_str().unwrap();
        let h = open(Level::Admin, p, OpenOptions::default()).unwrap();
        close(h).unwrap();
        assert!(close(h).is_err(), "二重 close がエラーにならなかった");
    }

    #[test]
    fn external_access_is_disabled_below_admin() {
        let path = tmp_db("external.db");
        let p = path.to_str().unwrap();
        let admin = seed(p);
        close(admin).unwrap();

        let h = open(Level::ReadWrite, p, OpenOptions::default()).unwrap();
        with_conn(h, |s| {
            // 外部ファイル読み取りが塞がれていること
            let err = s
                .conn
                .query("SELECT * FROM read_csv('C:/whatever.csv')")
                .unwrap_err();
            assert!(!err.is_empty(), "外部アクセスが拒否されなかった");
        })
        .unwrap();
        close(h).unwrap();
    }
}
