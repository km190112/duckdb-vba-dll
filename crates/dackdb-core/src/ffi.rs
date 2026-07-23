//! VBA から呼ばれる `#[no_mangle] extern "system"` 関数を生成するマクロ。
//!
//! # 呼び出し規約
//!
//! - **入力文字列**は `*const u16`。VBA 側は `ByVal p As LongPtr` と宣言して
//!   `StrPtr(s)` を渡す。VBA の `Declare` による ANSI マーシャリングを経由しないので
//!   日本語が化けない。
//! - **出力**は必ず最後の引数 `out: *mut VARIANT`。VBA 側は `ByRef result As Variant`。
//! - **戻り値**は `i32`。0 が成功、負値がエラー。**失敗時は `out` にエラーメッセージ
//!   （文字列）が入る**ので、VBA 側は一様に扱える。

/// FFI 境界の共通処理。
///
/// - `catch_unwind` で panic を捕まえる。**これが無いと Rust の panic が VBA へ
///   巻き戻って Excel がプロセスごと落ちる**。そのため release プロファイルで
///   `panic = "abort"` を指定してはいけない（指定すると catch_unwind が無力になる）。
/// - 成功／失敗のどちらでも `out` に書き込むので、VBA 側の分岐が 1 本で済む。
pub fn guard(
    out: *mut crate::oleaut::VARIANT,
    f: impl FnOnce() -> Result<crate::oleaut::VARIANT, String>,
) -> i32 {
    use crate::api::*;
    use crate::oleaut::VARIANT;
    use crate::variant::write_out;

    let result = std::panic::catch_unwind(std::panic::AssertUnwindSafe(f));

    match result {
        Ok(Ok(v)) => {
            unsafe { write_out(out, v) };
            DACK_OK
        }
        Ok(Err(msg)) => {
            let forbidden = msg.contains("この DLL では使えません");
            unsafe { write_out(out, VARIANT::bstr(&msg)) };
            if forbidden {
                DACK_E_FORBIDDEN
            } else {
                DACK_E_GENERAL
            }
        }
        Err(payload) => {
            // panic の内容をできる範囲で取り出して VBA に見せる。
            // 利用者が「Excel が落ちた」ではなく具体的な内容を報告できるようにする。
            let detail = if let Some(s) = payload.downcast_ref::<&str>() {
                (*s).to_string()
            } else if let Some(s) = payload.downcast_ref::<String>() {
                s.clone()
            } else {
                "詳細不明".to_string()
            };
            let msg = format!(
                "dackdb の内部エラー（panic）: {detail}\n\
                 これは DLL の不具合です。実行した SQL と合わせて報告してください。"
            );
            unsafe { write_out(out, VARIANT::bstr(&msg)) };
            DACK_E_PANIC
        }
    }
}

/// VBA から渡された UTF-16 ポインタを `String` にする。null なら分かりやすいエラー。
pub fn arg_string(p: *const u16, name: &str) -> Result<String, String> {
    unsafe { crate::oleaut::wide_ptr_to_string(p) }
        .ok_or_else(|| format!("引数 {name} が空（null）です。StrPtr() で渡してください。"))
}

/// 指定した権限レベルで全エクスポート関数を生成する。
///
/// 3 つの cdylib クレートは、それぞれこのマクロを 1 行呼ぶだけの殻になる。
///
/// # なぜ Cargo の feature ではなくマクロ引数なのか
///
/// 権限を feature flag で切り替える設計にすると、ワークスペースを一括ビルドした際に
/// **feature 統合（feature unification）** が起きて 3 つの DLL すべてが最も強い権限で
/// ビルドされてしまう。定数をマクロに渡せばクレートごとに独立する。
///
/// # 3 つの DLL が同じ関数名をエクスポートする理由
///
/// 権限外の関数も存在はし、`DACK_E_FORBIDDEN` と案内メッセージを返す。
/// こうすることで VBA 側モジュールは `Lib "..."` の文字列だけが違う 3 ファイルになり、
/// 階層② 向けに書いたコードが階層③ でそのまま動く。
#[macro_export]
macro_rules! export_dackdb_ffi {
    ($level:expr) => {
        /// この DLL の権限レベル。
        const DACK_LEVEL: $crate::Level = $level;

        type DackVariant = $crate::oleaut::VARIANT;

        // ---- 全階層 ----

        #[no_mangle]
        pub extern "system" fn DackVersion(out: *mut DackVariant) -> i32 {
            $crate::ffi::guard(out, || $crate::api::version(DACK_LEVEL))
        }

        #[no_mangle]
        pub extern "system" fn DackCapabilities(out: *mut DackVariant) -> i32 {
            $crate::ffi::guard(out, || $crate::api::capabilities(DACK_LEVEL))
        }

        #[no_mangle]
        pub extern "system" fn DackOpen(path: *const u16, out: *mut DackVariant) -> i32 {
            $crate::ffi::guard(out, || {
                let p = $crate::ffi::arg_string(path, "path")?;
                $crate::api::open(DACK_LEVEL, &p)
            })
        }

        #[no_mangle]
        pub extern "system" fn DackClose(handle: i64, out: *mut DackVariant) -> i32 {
            $crate::ffi::guard(out, || $crate::api::close(handle))
        }

        #[no_mangle]
        pub extern "system" fn DackQuery(
            handle: i64,
            sql: *const u16,
            out: *mut DackVariant,
        ) -> i32 {
            $crate::ffi::guard(out, || {
                let s = $crate::ffi::arg_string(sql, "sql")?;
                $crate::api::query(handle, &s)
            })
        }

        #[no_mangle]
        pub extern "system" fn DackQueryParams(
            handle: i64,
            sql: *const u16,
            params: *const DackVariant,
            out: *mut DackVariant,
        ) -> i32 {
            $crate::ffi::guard(out, || {
                let s = $crate::ffi::arg_string(sql, "sql")?;
                unsafe { $crate::api::query_params(handle, &s, params) }
            })
        }

        #[no_mangle]
        pub extern "system" fn DackListTables(handle: i64, out: *mut DackVariant) -> i32 {
            $crate::ffi::guard(out, || $crate::api::list_tables(handle))
        }

        #[no_mangle]
        pub extern "system" fn DackDescribe(
            handle: i64,
            table: *const u16,
            out: *mut DackVariant,
        ) -> i32 {
            $crate::ffi::guard(out, || {
                let t = $crate::ffi::arg_string(table, "table")?;
                $crate::api::describe(handle, &t)
            })
        }

        // ---- 階層② 以上（下位では DACK_E_FORBIDDEN） ----

        #[no_mangle]
        pub extern "system" fn DackExecute(
            handle: i64,
            sql: *const u16,
            out: *mut DackVariant,
        ) -> i32 {
            $crate::ffi::guard(out, || {
                let s = $crate::ffi::arg_string(sql, "sql")?;
                $crate::api::execute(DACK_LEVEL, handle, &s)
            })
        }

        #[no_mangle]
        pub extern "system" fn DackExecuteParams(
            handle: i64,
            sql: *const u16,
            params: *const DackVariant,
            out: *mut DackVariant,
        ) -> i32 {
            $crate::ffi::guard(out, || {
                let s = $crate::ffi::arg_string(sql, "sql")?;
                unsafe { $crate::api::execute_params(DACK_LEVEL, handle, &s, params) }
            })
        }

        #[no_mangle]
        pub extern "system" fn DackAppendArray(
            handle: i64,
            table: *const u16,
            data: *const DackVariant,
            out: *mut DackVariant,
        ) -> i32 {
            $crate::ffi::guard(out, || {
                let t = $crate::ffi::arg_string(table, "table")?;
                unsafe { $crate::api::append_array(DACK_LEVEL, handle, &t, data) }
            })
        }

        #[no_mangle]
        pub extern "system" fn DackBegin(handle: i64, out: *mut DackVariant) -> i32 {
            $crate::ffi::guard(out, || $crate::api::begin(DACK_LEVEL, handle))
        }

        #[no_mangle]
        pub extern "system" fn DackCommit(handle: i64, out: *mut DackVariant) -> i32 {
            $crate::ffi::guard(out, || $crate::api::commit(DACK_LEVEL, handle))
        }

        #[no_mangle]
        pub extern "system" fn DackRollback(handle: i64, out: *mut DackVariant) -> i32 {
            $crate::ffi::guard(out, || $crate::api::rollback(DACK_LEVEL, handle))
        }

        // ---- 階層③ のみ ----

        #[no_mangle]
        pub extern "system" fn DackCreateDatabase(
            path: *const u16,
            out: *mut DackVariant,
        ) -> i32 {
            $crate::ffi::guard(out, || {
                let p = $crate::ffi::arg_string(path, "path")?;
                $crate::api::create_database(DACK_LEVEL, &p)
            })
        }

        #[no_mangle]
        pub extern "system" fn DackExecuteDDL(
            handle: i64,
            sql: *const u16,
            out: *mut DackVariant,
        ) -> i32 {
            $crate::ffi::guard(out, || {
                let s = $crate::ffi::arg_string(sql, "sql")?;
                $crate::api::execute_ddl(DACK_LEVEL, handle, &s)
            })
        }

        #[no_mangle]
        pub extern "system" fn DackExportSchema(
            handle: i64,
            format: *const u16,
            out: *mut DackVariant,
        ) -> i32 {
            $crate::ffi::guard(out, || {
                let f = $crate::ffi::arg_string(format, "format")?;
                $crate::api::export_schema(DACK_LEVEL, handle, &f)
            })
        }

        #[no_mangle]
        pub extern "system" fn DackCheckpoint(handle: i64, out: *mut DackVariant) -> i32 {
            $crate::ffi::guard(out, || $crate::api::checkpoint(DACK_LEVEL, handle))
        }
    };
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::oleaut::*;

    #[test]
    fn guard_writes_value_and_returns_ok() {
        let mut out = VARIANT::empty();
        let rc = guard(&mut out, || Ok(VARIANT::i64(42)));
        assert_eq!(rc, crate::api::DACK_OK);
        assert_eq!(out.vt, VT_I8);
        assert_eq!(unsafe { out.value.llVal }, 42);
        out.clear();
    }

    #[test]
    fn guard_writes_error_message_into_the_same_out_param() {
        let mut out = VARIANT::empty();
        let rc = guard(&mut out, || Err("テーブルが見つかりません".to_string()));
        assert_eq!(rc, crate::api::DACK_E_GENERAL);
        assert_eq!(out.vt, VT_BSTR);
        assert_eq!(
            unsafe { bstr_to_string(out.value.bstrVal) },
            "テーブルが見つかりません"
        );
        out.clear();
    }

    /// panic が VBA へ巻き戻ると Excel がプロセスごと落ちる。
    /// ここで捕まえられていることが最重要。
    #[test]
    fn guard_catches_panic_instead_of_unwinding_into_vba() {
        let prev = std::panic::take_hook();
        std::panic::set_hook(Box::new(|_| {})); // テスト出力を汚さない
        let mut out = VARIANT::empty();
        let rc = guard(&mut out, || panic!("意図的なテスト用 panic"));
        std::panic::set_hook(prev);

        assert_eq!(rc, crate::api::DACK_E_PANIC);
        assert_eq!(out.vt, VT_BSTR);
        let msg = unsafe { bstr_to_string(out.value.bstrVal) };
        assert!(msg.contains("意図的なテスト用 panic"), "panic 内容が失われた: {msg}");
        assert!(msg.contains("報告"), "報告の案内が無い: {msg}");
        out.clear();
    }

    #[test]
    fn guard_reports_forbidden_separately_from_general_errors() {
        let mut out = VARIANT::empty();
        let rc = guard(&mut out, || {
            Err("DackExecute はこの DLL では使えません。".to_string())
        });
        assert_eq!(rc, crate::api::DACK_E_FORBIDDEN);
        out.clear();
    }

    #[test]
    fn arg_string_rejects_null_with_guidance() {
        let err = arg_string(std::ptr::null(), "sql").unwrap_err();
        assert!(err.contains("StrPtr"), "{err}");
    }

    #[test]
    fn arg_string_reads_japanese_utf16() {
        let mut w: Vec<u16> = "SELECT * FROM 売上".encode_utf16().collect();
        w.push(0);
        assert_eq!(arg_string(w.as_ptr(), "sql").unwrap(), "SELECT * FROM 売上");
    }
}
