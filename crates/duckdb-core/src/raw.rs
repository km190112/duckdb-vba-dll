//! `libduckdb-sys`（DuckDB の生 C API）に対する薄い RAII ラッパ。
//!
//! `duckdb-rs`（安全ラッパ）ではなく生 C API を使っている理由：
//! `duckdb_prepared_statement_type` と `duckdb_extract_statements` が必要で、
//! これらは `duckdb-rs` の公開 API に無い。詳細は `classify.rs` 参照。

use libduckdb_sys as ffi;
use std::ffi::{CStr, CString};
use std::os::raw::c_char;

/// C API に渡すための NUL 終端文字列を作る。
///
/// SQL 文中に NUL が含まれていたら（VBA からは通常起こらないが）エラーにする。
/// ここで黙って切り捨てると、検査した SQL と実行する SQL がズレて権限ゲートを
/// すり抜ける余地が生まれるため、必ず拒否する。
pub fn to_cstring(s: &str) -> Result<CString, String> {
    CString::new(s).map_err(|_| "文字列に NUL 文字が含まれています".to_string())
}

/// `*const c_char` を Rust の `String` に（null 安全）。
unsafe fn cstr_to_string(p: *const c_char) -> String {
    if p.is_null() {
        String::new()
    } else {
        CStr::from_ptr(p).to_string_lossy().into_owned()
    }
}

// ---------------------------------------------------------------------------
// Config
// ---------------------------------------------------------------------------

/// `duckdb_config` の RAII ラッパ。
pub struct Config {
    handle: ffi::duckdb_config,
}

impl Config {
    pub fn new() -> Result<Self, String> {
        let mut handle: ffi::duckdb_config = std::ptr::null_mut();
        let state = unsafe { ffi::duckdb_create_config(&mut handle) };
        if state != ffi::DuckDBSuccess || handle.is_null() {
            return Err("DuckDB の設定オブジェクトを作成できませんでした".to_string());
        }
        Ok(Config { handle })
    }

    /// 設定を 1 つ書き込む。
    ///
    /// 注意：`lock_configuration` は**必ず最後に**設定すること。先に設定すると
    /// 以降の `duckdb_set_config` が効かなくなる。
    pub fn set(&mut self, name: &str, value: &str) -> Result<(), String> {
        let cname = to_cstring(name)?;
        let cvalue = to_cstring(value)?;
        let state = unsafe { ffi::duckdb_set_config(self.handle, cname.as_ptr(), cvalue.as_ptr()) };
        if state != ffi::DuckDBSuccess {
            return Err(format!("DuckDB の設定 '{name}' = '{value}' に失敗しました"));
        }
        Ok(())
    }
}

impl Drop for Config {
    fn drop(&mut self) {
        if !self.handle.is_null() {
            unsafe { ffi::duckdb_destroy_config(&mut self.handle) };
        }
    }
}

// ---------------------------------------------------------------------------
// Database / Connection
// ---------------------------------------------------------------------------

/// `duckdb_database` の RAII ラッパ。
pub struct Database {
    handle: ffi::duckdb_database,
}

// ハンドルはグローバルなレジストリ（Mutex 保護）に格納するため Send が要る。
// アクセスは必ず Mutex 経由で直列化される。
unsafe impl Send for Database {}

impl Database {
    /// 設定付きでデータベースを開く。
    pub fn open(path: &str, config: &Config) -> Result<Self, String> {
        let cpath = to_cstring(path)?;
        let mut handle: ffi::duckdb_database = std::ptr::null_mut();
        let mut err: *mut c_char = std::ptr::null_mut();

        let state =
            unsafe { ffi::duckdb_open_ext(cpath.as_ptr(), &mut handle, config.handle, &mut err) };

        if state != ffi::DuckDBSuccess {
            let msg = unsafe { cstr_to_string(err) };
            if !err.is_null() {
                unsafe { ffi::duckdb_free(err as *mut _) };
            }
            return Err(if msg.is_empty() {
                format!("データベース '{path}' を開けませんでした")
            } else {
                format!("データベース '{path}' を開けませんでした: {msg}")
            });
        }
        // 成功時もエラーバッファが確保されている可能性があるので解放しておく
        if !err.is_null() {
            unsafe { ffi::duckdb_free(err as *mut _) };
        }
        Ok(Database { handle })
    }

    pub fn connect(&self) -> Result<Connection, String> {
        let mut handle: ffi::duckdb_connection = std::ptr::null_mut();
        let state = unsafe { ffi::duckdb_connect(self.handle, &mut handle) };
        if state != ffi::DuckDBSuccess || handle.is_null() {
            return Err("データベースへの接続に失敗しました".to_string());
        }
        Ok(Connection { handle })
    }
}

impl Drop for Database {
    fn drop(&mut self) {
        if !self.handle.is_null() {
            unsafe { ffi::duckdb_close(&mut self.handle) };
        }
    }
}

/// `duckdb_connection` の RAII ラッパ。
pub struct Connection {
    handle: ffi::duckdb_connection,
}

unsafe impl Send for Connection {}

impl Connection {
    /// 生ハンドル。`classify.rs` が `duckdb_extract_statements` に渡すために使う。
    pub fn raw(&self) -> ffi::duckdb_connection {
        self.handle
    }

    /// SQL を実行して結果を得る。**権限検査は呼び出し側の責任**。
    pub fn query(&self, sql: &str) -> Result<QueryResult, String> {
        let csql = to_cstring(sql)?;
        let mut raw: ffi::duckdb_result = unsafe { std::mem::zeroed() };
        let state = unsafe { ffi::duckdb_query(self.handle, csql.as_ptr(), &mut raw) };
        let mut result = QueryResult { raw };
        if state != ffi::DuckDBSuccess {
            let msg = unsafe { cstr_to_string(ffi::duckdb_result_error(&mut result.raw)) };
            return Err(if msg.is_empty() {
                "SQL の実行に失敗しました".to_string()
            } else {
                msg
            });
        }
        Ok(result)
    }
}

impl Drop for Connection {
    fn drop(&mut self) {
        if !self.handle.is_null() {
            unsafe { ffi::duckdb_disconnect(&mut self.handle) };
        }
    }
}

// ---------------------------------------------------------------------------
// QueryResult
// ---------------------------------------------------------------------------

/// `duckdb_result` の RAII ラッパ。
pub struct QueryResult {
    raw: ffi::duckdb_result,
}

impl QueryResult {
    pub fn column_count(&mut self) -> usize {
        unsafe { ffi::duckdb_column_count(&mut self.raw) as usize }
    }

    pub fn column_name(&mut self, col: usize) -> String {
        unsafe { cstr_to_string(ffi::duckdb_column_name(&mut self.raw, col as ffi::idx_t)) }
    }

    /// 列の論理型 ID。データチャンクを取り出す前に取得できるので、
    /// 変換できない型を先に検出してクエリごと拒否するのに使う。
    pub fn column_type(&mut self, col: usize) -> ffi::duckdb_type {
        unsafe { ffi::duckdb_column_type(&mut self.raw, col as ffi::idx_t) }
    }

    /// INSERT / UPDATE / DELETE が変更した行数。
    pub fn rows_changed(&mut self) -> i64 {
        unsafe { ffi::duckdb_rows_changed(&mut self.raw) as i64 }
    }

    /// 次のデータチャンクを取り出す。`None` なら終端。
    pub fn next_chunk(&mut self) -> Option<DataChunk> {
        let chunk = unsafe { ffi::duckdb_fetch_chunk(self.raw) };
        if chunk.is_null() {
            None
        } else {
            Some(DataChunk { handle: chunk })
        }
    }
}

impl Drop for QueryResult {
    fn drop(&mut self) {
        unsafe { ffi::duckdb_destroy_result(&mut self.raw) };
    }
}

/// `Result<QueryResult, String>` に対して `unwrap_err()` を使えるようにするため。
impl std::fmt::Debug for QueryResult {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        f.write_str("QueryResult")
    }
}

/// `duckdb_data_chunk` の RAII ラッパ。
pub struct DataChunk {
    handle: ffi::duckdb_data_chunk,
}

impl DataChunk {
    /// このチャンクに入っている行数。
    pub fn len(&self) -> usize {
        unsafe { ffi::duckdb_data_chunk_get_size(self.handle) as usize }
    }

    pub fn is_empty(&self) -> bool {
        self.len() == 0
    }

    pub fn column_count(&self) -> usize {
        unsafe { ffi::duckdb_data_chunk_get_column_count(self.handle) as usize }
    }

    /// 指定列のベクタ。返り値はチャンクの生存期間に束縛される。
    pub fn vector(&self, col: usize) -> Vector<'_> {
        let v = unsafe { ffi::duckdb_data_chunk_get_vector(self.handle, col as ffi::idx_t) };
        Vector {
            handle: v,
            _chunk: std::marker::PhantomData,
        }
    }
}

impl Drop for DataChunk {
    fn drop(&mut self) {
        if !self.handle.is_null() {
            unsafe { ffi::duckdb_destroy_data_chunk(&mut self.handle) };
        }
    }
}

/// `duckdb_vector`。所有権を持たず、親チャンクの生存期間に束縛される。
pub struct Vector<'a> {
    handle: ffi::duckdb_vector,
    _chunk: std::marker::PhantomData<&'a DataChunk>,
}

impl Vector<'_> {
    /// この列の論理型 ID（`DUCKDB_TYPE_*`）。
    pub fn type_id(&self) -> ffi::duckdb_type {
        unsafe {
            let lt = ffi::duckdb_vector_get_column_type(self.handle);
            let id = ffi::duckdb_get_type_id(lt);
            let mut lt_mut = lt;
            ffi::duckdb_destroy_logical_type(&mut lt_mut);
            id
        }
    }

    /// DECIMAL の (width, scale)。DECIMAL 以外では意味を持たない。
    pub fn decimal_width_scale(&self) -> (u8, u8) {
        unsafe {
            let lt = ffi::duckdb_vector_get_column_type(self.handle);
            let w = ffi::duckdb_decimal_width(lt);
            let s = ffi::duckdb_decimal_scale(lt);
            let mut lt_mut = lt;
            ffi::duckdb_destroy_logical_type(&mut lt_mut);
            (w, s)
        }
    }

    /// DECIMAL の内部表現型（SMALLINT / INTEGER / BIGINT / HUGEINT のいずれか）。
    pub fn decimal_internal_type(&self) -> ffi::duckdb_type {
        unsafe {
            let lt = ffi::duckdb_vector_get_column_type(self.handle);
            let t = ffi::duckdb_decimal_internal_type(lt);
            let mut lt_mut = lt;
            ffi::duckdb_destroy_logical_type(&mut lt_mut);
            t
        }
    }

    /// 生の論理型ハンドル。**呼び出し側が `duckdb_destroy_logical_type` すること**。
    /// ENUM の辞書引きのように、型 ID だけでは足りない場合に使う。
    pub fn logical_type_raw(&self) -> ffi::duckdb_logical_type {
        unsafe { ffi::duckdb_vector_get_column_type(self.handle) }
    }

    /// 生データ配列の先頭。要素サイズは型ごとに異なる。
    pub fn data_ptr(&self) -> *mut std::ffi::c_void {
        unsafe { ffi::duckdb_vector_get_data(self.handle) }
    }

    /// 行 `row` が NULL でないかを返す。
    ///
    /// validity マスクが null の場合、その列に NULL は一切無い（DuckDB の規約）。
    pub fn is_valid(&self, row: usize) -> bool {
        unsafe {
            let mask = ffi::duckdb_vector_get_validity(self.handle);
            if mask.is_null() {
                return true;
            }
            ffi::duckdb_validity_row_is_valid(mask, row as ffi::idx_t)
        }
    }
}

// ---------------------------------------------------------------------------
// Prepared / Extracted（権限判定に使う）
// ---------------------------------------------------------------------------

/// `duckdb_extracted_statements` の RAII ラッパ。
pub struct Extracted {
    pub(crate) handle: ffi::duckdb_extracted_statements,
    pub(crate) count: usize,
}

impl Extracted {
    /// SQL 文字列を DuckDB のパーサで個々の文に分割する。
    ///
    /// 自前の正規表現ではなく DuckDB 自身のパーサを使うのが要点。
    /// `SELECT 1; DROP TABLE t;` のような複数文の混入を確実に検出できる。
    pub fn extract(conn: &Connection, sql: &str) -> Result<Self, String> {
        let csql = to_cstring(sql)?;
        let mut handle: ffi::duckdb_extracted_statements = std::ptr::null_mut();
        let count =
            unsafe { ffi::duckdb_extract_statements(conn.raw(), csql.as_ptr(), &mut handle) }
                as usize;

        if count == 0 {
            let msg = unsafe { cstr_to_string(ffi::duckdb_extract_statements_error(handle)) };
            unsafe { ffi::duckdb_destroy_extracted(&mut handle) };
            return Err(if msg.is_empty() {
                "SQL を解析できませんでした（空の文の可能性があります）".to_string()
            } else {
                format!("SQL の構文エラー: {msg}")
            });
        }
        Ok(Extracted { handle, count })
    }

    pub fn count(&self) -> usize {
        self.count
    }

    /// `index` 番目の文を prepare する。型判定のために必要。
    pub fn prepare(&self, conn: &Connection, index: usize) -> Result<Prepared, String> {
        let mut stmt: ffi::duckdb_prepared_statement = std::ptr::null_mut();
        let state = unsafe {
            ffi::duckdb_prepare_extracted_statement(
                conn.raw(),
                self.handle,
                index as ffi::idx_t,
                &mut stmt,
            )
        };
        let prepared = Prepared { handle: stmt };
        if state != ffi::DuckDBSuccess {
            let msg = unsafe { cstr_to_string(ffi::duckdb_prepare_error(prepared.handle)) };
            return Err(if msg.is_empty() {
                format!("{} 文目を解析できませんでした", index + 1)
            } else {
                msg
            });
        }
        Ok(prepared)
    }
}

impl Drop for Extracted {
    fn drop(&mut self) {
        if !self.handle.is_null() {
            unsafe { ffi::duckdb_destroy_extracted(&mut self.handle) };
        }
    }
}

/// `duckdb_prepared_statement` の RAII ラッパ。
pub struct Prepared {
    handle: ffi::duckdb_prepared_statement,
}

impl Prepared {
    /// DuckDB 自身が判定した文の種別。権限ゲートの根拠になる。
    pub fn statement_type(&self) -> ffi::duckdb_statement_type {
        unsafe { ffi::duckdb_prepared_statement_type(self.handle) }
    }

    pub fn param_count(&self) -> usize {
        unsafe { ffi::duckdb_nparams(self.handle) as usize }
    }

    /// `?` の `index` 番目（1 始まり）に値をバインドする。
    ///
    /// # Safety
    /// `value` は有効な `duckdb_value` であること（呼び出し側が所有権を保持する）。
    pub unsafe fn bind(&self, index: usize, value: ffi::duckdb_value) -> Result<(), String> {
        let state = unsafe { ffi::duckdb_bind_value(self.handle, index as ffi::idx_t, value) };
        if state != ffi::DuckDBSuccess {
            let msg = unsafe { cstr_to_string(ffi::duckdb_prepare_error(self.handle)) };
            return Err(if msg.is_empty() {
                format!("{index} 番目のパラメータをバインドできませんでした")
            } else {
                format!("{index} 番目のパラメータ: {msg}")
            });
        }
        Ok(())
    }

    /// バインド済みの文を実行する。
    pub fn execute(&self) -> Result<QueryResult, String> {
        let mut raw: ffi::duckdb_result = unsafe { std::mem::zeroed() };
        let state = unsafe { ffi::duckdb_execute_prepared(self.handle, &mut raw) };
        let mut result = QueryResult { raw };
        if state != ffi::DuckDBSuccess {
            let msg = unsafe { cstr_to_string(ffi::duckdb_result_error(&mut result.raw)) };
            return Err(if msg.is_empty() {
                "SQL の実行に失敗しました".to_string()
            } else {
                msg
            });
        }
        Ok(result)
    }
}

// ---------------------------------------------------------------------------
// Appender（一括投入）
// ---------------------------------------------------------------------------

/// `duckdb_appender` の RAII ラッパ。
///
/// INSERT 文を 1 行ずつ実行するより 1〜2 桁速い。Excel のシート範囲を
/// そのままテーブルへ流し込む用途のための経路。
pub struct Appender {
    handle: ffi::duckdb_appender,
}

impl Appender {
    pub fn new(conn: &Connection, table: &str) -> Result<Self, String> {
        let ctable = to_cstring(table)?;
        let mut handle: ffi::duckdb_appender = std::ptr::null_mut();
        let state = unsafe {
            // schema に null を渡すと既定のスキーマが使われる
            ffi::duckdb_appender_create(conn.raw(), std::ptr::null(), ctable.as_ptr(), &mut handle)
        };
        let appender = Appender { handle };
        if state != ffi::DuckDBSuccess {
            let msg = appender.error();
            return Err(if msg.is_empty() {
                format!("テーブル「{table}」を開けませんでした。名前を確認してください。")
            } else {
                msg
            });
        }
        Ok(appender)
    }

    /// 投入先テーブルの列数。渡された配列の列数と突き合わせるのに使う。
    pub fn column_count(&self) -> usize {
        unsafe { ffi::duckdb_appender_column_count(self.handle) as usize }
    }

    /// # Safety
    /// `value` は有効な `duckdb_value` であること（呼び出し側が所有権を保持する）。
    pub unsafe fn append_value(&self, value: ffi::duckdb_value) -> Result<(), String> {
        if ffi::duckdb_append_value(self.handle, value) != ffi::DuckDBSuccess {
            return Err(self.error_or("値を追加できませんでした"));
        }
        Ok(())
    }

    pub fn end_row(&self) -> Result<(), String> {
        if unsafe { ffi::duckdb_appender_end_row(self.handle) } != ffi::DuckDBSuccess {
            return Err(self.error_or("行を確定できませんでした"));
        }
        Ok(())
    }

    /// バッファをテーブルへ書き出す。**成功を確認するには必ず呼ぶこと**
    /// （Drop 任せにすると失敗を検出できない）。
    pub fn flush(&self) -> Result<(), String> {
        if unsafe { ffi::duckdb_appender_flush(self.handle) } != ffi::DuckDBSuccess {
            return Err(self.error_or("テーブルへの書き出しに失敗しました"));
        }
        Ok(())
    }

    fn error(&self) -> String {
        unsafe { cstr_to_string(ffi::duckdb_appender_error(self.handle)) }
    }

    fn error_or(&self, fallback: &str) -> String {
        let msg = self.error();
        if msg.is_empty() {
            fallback.to_string()
        } else {
            msg
        }
    }
}

impl Drop for Appender {
    fn drop(&mut self) {
        if !self.handle.is_null() {
            unsafe { ffi::duckdb_appender_destroy(&mut self.handle) };
        }
    }
}

impl Drop for Prepared {
    fn drop(&mut self) {
        if !self.handle.is_null() {
            unsafe { ffi::duckdb_destroy_prepare(&mut self.handle) };
        }
    }
}
