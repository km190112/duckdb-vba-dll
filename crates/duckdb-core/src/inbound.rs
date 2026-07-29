//! VBA から**入力として**渡される VARIANT / SAFEARRAY の読み取り。
//!
//! 出力側（[`crate::variant`]）と対になるが、処理はまったく別物になる。
//!
//! - 出力側は自分で確保した SAFEARRAY に書き込み、所有権を VBA に渡す。
//! - 入力側は **VBA が所有するメモリを読むだけ**。ここで受け取った BSTR や
//!   SAFEARRAY を解放してはいけない（VBA 側が二重解放でクラッシュする）。

#![allow(non_upper_case_globals)]

use crate::oleaut::*;
use libduckdb_sys as ffi;
use std::ffi::c_void;

/// OLE 日付（1899-12-30 起点の日数）→ Unix エポックからのマイクロ秒。
const OLE_EPOCH_OFFSET_DAYS: f64 = 25569.0;
const MICROS_PER_DAY: f64 = 86_400_000_000.0;

// ---------------------------------------------------------------------------
// duckdb_value の RAII ラッパ
// ---------------------------------------------------------------------------

/// `duckdb_value` の RAII ラッパ。バインドと Appender の両方で使う。
pub struct DuckValue {
    raw: ffi::duckdb_value,
}

impl DuckValue {
    fn from_raw(raw: ffi::duckdb_value, what: &str) -> Result<Self, String> {
        if raw.is_null() {
            return Err(format!("{what} を DuckDB の値に変換できませんでした"));
        }
        Ok(DuckValue { raw })
    }

    pub fn raw(&self) -> ffi::duckdb_value {
        self.raw
    }
}

impl Drop for DuckValue {
    fn drop(&mut self) {
        if !self.raw.is_null() {
            unsafe { ffi::duckdb_destroy_value(&mut self.raw) };
        }
    }
}

/// `Result<DuckValue, String>` に `unwrap_err()` を使えるようにするため。
impl std::fmt::Debug for DuckValue {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        f.write_str("DuckValue")
    }
}

// ---------------------------------------------------------------------------
// VARIANT → duckdb_value
// ---------------------------------------------------------------------------

/// VBA から来た 1 つの VARIANT を DuckDB の値に変換する。
///
/// `where_` はエラーメッセージ用の位置情報（例: `"3 行 2 列目"`）。
///
/// # Safety
/// `v` は有効な VARIANT であること。中身の BSTR / SAFEARRAY は解放しない。
pub unsafe fn variant_to_value(v: &VARIANT, where_: &str) -> Result<DuckValue, String> {
    // VBA が ByRef で渡してくる場合に備えて VT_BYREF を辿る。
    if v.vt & VT_BYREF != 0 {
        let inner = v.value.byref;
        if inner.is_null() {
            return DuckValue::from_raw(ffi::duckdb_create_null_value(), where_);
        }
        // VT_BYREF|VT_VARIANT なら中身は VARIANT。それ以外は基本型へのポインタ。
        if v.vt & VT_TYPEMASK == VT_VARIANT {
            return variant_to_value(&*(inner as *const VARIANT), where_);
        }
        let mut deref = VARIANT::empty();
        deref.vt = v.vt & !VT_BYREF;
        match deref.vt {
            VT_I4 => deref.value.lVal = *(inner as *const i32),
            VT_I8 => deref.value.llVal = *(inner as *const i64),
            VT_R8 | VT_DATE => deref.value.dblVal = *(inner as *const f64),
            VT_BOOL => deref.value.boolVal = *(inner as *const i16),
            VT_BSTR => deref.value.bstrVal = *(inner as *const *mut u16),
            _ => {
                return Err(format!(
                    "{where_}: 対応していない参照渡しの型 (vt=0x{:04X})",
                    v.vt
                ))
            }
        }
        return variant_to_value(&deref, where_);
    }

    // バイト配列は BLOB として扱う（他の配列は非対応）。
    if v.vt & VT_ARRAY != 0 {
        if v.vt & VT_TYPEMASK == VT_UI1 {
            return byte_array_to_blob(v.value.parray, where_);
        }
        return Err(format!(
            "{where_}: 配列は値として渡せません。1 つのセルの値を指定してください。"
        ));
    }

    let raw = match v.vt {
        VT_EMPTY | VT_NULL => ffi::duckdb_create_null_value(),

        VT_BOOL => ffi::duckdb_create_bool(v.value.boolVal != VARIANT_FALSE),

        VT_I1 => ffi::duckdb_create_int64(*(&v.value.bVal as *const u8 as *const i8) as i64),
        VT_I2 => ffi::duckdb_create_int64(v.value.iVal as i64),
        VT_I4 | VT_INT => ffi::duckdb_create_int64(v.value.lVal as i64),
        VT_I8 => ffi::duckdb_create_int64(v.value.llVal),
        VT_UI1 => ffi::duckdb_create_int64(v.value.bVal as i64),
        VT_UI2 => ffi::duckdb_create_int64((v.value.iVal as u16) as i64),
        VT_UI4 | VT_UINT => ffi::duckdb_create_int64((v.value.lVal as u32) as i64),
        VT_UI8 => {
            let u = v.value.llVal as u64;
            if u > i64::MAX as u64 {
                // i64 に収まらない値は精度を落とさず文字列で渡し、DuckDB にキャストさせる
                return varchar_value(&u.to_string(), where_);
            }
            ffi::duckdb_create_int64(u as i64)
        }

        VT_R4 => ffi::duckdb_create_double(v.value.fltVal as f64),
        VT_R8 => ffi::duckdb_create_double(v.value.dblVal),

        // Excel のセルの日付は OLE 日付。DuckDB には TIMESTAMP として渡し、
        // 列が DATE なら DuckDB 側でキャストされる。
        VT_DATE => {
            let micros = ((v.value.date - OLE_EPOCH_OFFSET_DAYS) * MICROS_PER_DAY).round();
            if !micros.is_finite() {
                return Err(format!("{where_}: 日付の値が不正です"));
            }
            ffi::duckdb_create_timestamp(ffi::duckdb_timestamp { micros: micros as i64 })
        }

        VT_BSTR => {
            let s = bstr_to_string(v.value.bstrVal);
            return varchar_value(&s, where_);
        }

        // CURRENCY と DECIMAL は OLE に変換を任せる。
        VT_CY | VT_DECIMAL => {
            let d = change_to_f64(v).ok_or_else(|| {
                format!("{where_}: 数値に変換できませんでした (vt=0x{:04X})", v.vt)
            })?;
            ffi::duckdb_create_double(d)
        }

        // Excel の #N/A などのエラー値。NULL に落とすとデータが静かに壊れるので拒否する。
        VT_ERROR => {
            return Err(format!(
                "{where_}: セルがエラー値（#N/A、#DIV/0! など）です。\n\
                 修正するか、IFERROR で空欄に置き換えてから渡してください。"
            ))
        }

        VT_DISPATCH | VT_UNKNOWN => {
            return Err(format!(
                "{where_}: オブジェクトは値として渡せません。Range ではなく Range.Value を渡してください。"
            ))
        }

        other => return Err(format!("{where_}: 対応していない型です (vt=0x{other:04X})")),
    };

    DuckValue::from_raw(raw, where_)
}

unsafe fn varchar_value(s: &str, where_: &str) -> Result<DuckValue, String> {
    let raw = ffi::duckdb_create_varchar_length(
        s.as_ptr() as *const std::os::raw::c_char,
        s.len() as ffi::idx_t,
    );
    DuckValue::from_raw(raw, where_)
}

unsafe fn byte_array_to_blob(psa: *mut SAFEARRAY, where_: &str) -> Result<DuckValue, String> {
    if psa.is_null() {
        return DuckValue::from_raw(ffi::duckdb_create_null_value(), where_);
    }
    let mut lb = 0i32;
    let mut ub = 0i32;
    if SafeArrayGetLBound(psa, 1, &mut lb) != S_OK || SafeArrayGetUBound(psa, 1, &mut ub) != S_OK {
        return Err(format!("{where_}: バイト配列の範囲を取得できませんでした"));
    }
    let len = (ub - lb + 1).max(0) as usize;
    let mut data: *mut c_void = std::ptr::null_mut();
    if SafeArrayAccessData(psa, &mut data) != S_OK || data.is_null() {
        return Err(format!("{where_}: バイト配列を読み取れませんでした"));
    }
    let raw = ffi::duckdb_create_blob(data as *const u8, len as ffi::idx_t);
    SafeArrayUnaccessData(psa);
    DuckValue::from_raw(raw, where_)
}

/// OLE の変換機能で f64 にする。
unsafe fn change_to_f64(v: &VARIANT) -> Option<f64> {
    let mut dst = VARIANT::empty();
    VariantInit(&mut dst);
    if VariantChangeType(&mut dst, v as *const VARIANT, 0, VT_R8) != S_OK {
        return None;
    }
    let d = dst.value.dblVal;
    dst.clear();
    Some(d)
}

// ---------------------------------------------------------------------------
// 入力 SAFEARRAY の読み取り
// ---------------------------------------------------------------------------

/// VBA から渡された配列の中身。要素は**借用**であり解放しない。
pub struct InputGrid {
    psa: *mut SAFEARRAY,
    pub rows: usize,
    pub cols: usize,
    /// 行優先に並べ替えた各要素のコピー（VARIANT は Copy な POD）。
    cells: Vec<VARIANT>,
}

impl InputGrid {
    /// `row`（0 始まり）, `col`（0 始まり）の要素。
    pub fn get(&self, row: usize, col: usize) -> &VARIANT {
        &self.cells[row * self.cols + col]
    }

    /// 人間向けの位置表記（エラーメッセージ用、1 始まり）。
    pub fn position(&self, row: usize, col: usize) -> String {
        format!("{} 行 {} 列目", row + 1, col + 1)
    }
}

/// `Result<InputGrid, String>` に `unwrap_err()` を使えるようにするため。
impl std::fmt::Debug for InputGrid {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(f, "InputGrid({} 行 x {} 列)", self.rows, self.cols)
    }
}

impl Drop for InputGrid {
    fn drop(&mut self) {
        // AccessData のロックを必ず解除する。中身は VBA の所有物なので解放しない。
        if !self.psa.is_null() {
            unsafe { SafeArrayUnaccessData(self.psa) };
        }
    }
}

/// 1 次元または 2 次元の Variant 配列を読み取る。
///
/// 1 次元の場合は「1 行 N 列」として扱う。VBA の `Array(1, 2, 3)` と
/// `Range("A1:C1").Value` を同じように扱えるようにするため。
///
/// # Safety
/// `v` は VBA が所有する有効な VARIANT であること。
pub unsafe fn read_input_grid(v: &VARIANT, what: &str) -> Result<InputGrid, String> {
    // ByRef で包まれている場合を辿る
    if v.vt & VT_BYREF != 0 && v.vt & VT_TYPEMASK == VT_VARIANT {
        let inner = v.value.byref;
        if inner.is_null() {
            return Err(format!("{what} が空です。"));
        }
        return read_input_grid(&*(inner as *const VARIANT), what);
    }

    if v.vt & VT_ARRAY == 0 {
        return Err(format!(
            "{what} は配列ではありません。\n\
             VBA の Array(...) かシートの範囲（Range(\"A1:C10\").Value）を渡してください。"
        ));
    }

    let psa = v.value.parray;
    if psa.is_null() {
        return Err(format!("{what} が空の配列です。"));
    }

    // 要素が Variant であることを確認する。Excel の Range.Value は必ず Variant。
    let mut elem_vt: u16 = 0;
    if SafeArrayGetVartype(psa, &mut elem_vt) != S_OK {
        return Err(format!("{what} の要素型を取得できませんでした。"));
    }
    if elem_vt != VT_VARIANT {
        return Err(format!(
            "{what} の要素が Variant ではありません (vt=0x{elem_vt:04X})。\n\
             Dim arr() As Variant で宣言するか、Range(...).Value を渡してください。"
        ));
    }

    let dims = SafeArrayGetDim(psa);
    if dims == 0 || dims > 2 {
        return Err(format!(
            "{what} は 1 次元か 2 次元の配列である必要があります（実際: {dims} 次元）。"
        ));
    }

    let bound = |d: u32| -> Result<(i32, i32), String> {
        let mut lb = 0i32;
        let mut ub = 0i32;
        if SafeArrayGetLBound(psa, d, &mut lb) != S_OK
            || SafeArrayGetUBound(psa, d, &mut ub) != S_OK
        {
            return Err(format!("{what} の次元 {d} の範囲を取得できませんでした。"));
        }
        Ok((lb, ub))
    };

    let (rows, cols, lb1, lb2) = if dims == 1 {
        let (lb, ub) = bound(1)?;
        (1usize, (ub - lb + 1).max(0) as usize, lb, 0)
    } else {
        let (lb1, ub1) = bound(1)?;
        let (lb2, ub2) = bound(2)?;
        (
            (ub1 - lb1 + 1).max(0) as usize,
            (ub2 - lb2 + 1).max(0) as usize,
            lb1,
            lb2,
        )
    };

    if rows == 0 || cols == 0 {
        return Err(format!("{what} が空です。"));
    }

    let _ = (lb1, lb2); // 下限は AccessData での線形位置計算には影響しない

    let mut data: *mut c_void = std::ptr::null_mut();
    if SafeArrayAccessData(psa, &mut data) != S_OK || data.is_null() {
        return Err(format!("{what} の内容を読み取れませんでした。"));
    }

    // SAFEARRAY は列優先（Fortran 順）。要素 (r, c) は c * rows + r の位置。
    // 出力側 variant.rs と同じ規則。
    let src = data as *const VARIANT;
    let mut cells = Vec::with_capacity(rows * cols);
    for r in 0..rows {
        for c in 0..cols {
            let linear = if dims == 1 { c } else { c * rows + r };
            cells.push(*src.add(linear));
        }
    }

    Ok(InputGrid {
        psa,
        rows,
        cols,
        cells,
    })
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::variant::Grid;

    /// 出力側の Grid を作り、それを入力として読み直す往復テスト。
    /// 列優先インデックスの解釈が出力側と入力側で一致していることの証明になる。
    fn roundtrip(rows: Vec<Vec<VARIANT>>, cols: usize) -> (VARIANT, InputGrid) {
        let mut g = Grid::with_cols(cols);
        for r in rows {
            g.push_row(r);
        }
        let v = g.into_variant().expect("配列化に失敗");
        let grid = unsafe { read_input_grid(&v, "テストデータ") }.expect("読み取りに失敗");
        (v, grid)
    }

    #[test]
    fn two_dimensional_array_roundtrips_with_matching_orientation() {
        let (mut v, grid) = roundtrip(
            vec![
                vec![VARIANT::i32(11), VARIANT::i32(12), VARIANT::i32(13)],
                vec![VARIANT::i32(21), VARIANT::i32(22), VARIANT::i32(23)],
            ],
            3,
        );
        assert_eq!((grid.rows, grid.cols), (2, 3));
        unsafe {
            assert_eq!(grid.get(0, 0).value.lVal, 11);
            assert_eq!(grid.get(0, 2).value.lVal, 13);
            assert_eq!(grid.get(1, 0).value.lVal, 21);
            assert_eq!(grid.get(1, 2).value.lVal, 23, "行と列が入れ替わっている");
        }
        drop(grid);
        v.clear();
    }

    #[test]
    fn japanese_strings_survive_the_input_path() {
        let (mut v, grid) = roundtrip(vec![vec![VARIANT::bstr("技術部𠮷😀")]], 1);
        unsafe {
            let s = bstr_to_string(grid.get(0, 0).value.bstrVal);
            assert_eq!(s, "技術部𠮷😀");
        }
        drop(grid);
        v.clear();
    }

    #[test]
    fn non_array_input_is_rejected_with_guidance() {
        let v = VARIANT::i64(5);
        let err = unsafe { read_input_grid(&v, "パラメータ") }.unwrap_err();
        assert!(err.contains("配列ではありません"), "{err}");
        assert!(err.contains("Range"), "対処方法が示されていない: {err}");
    }

    #[test]
    fn scalar_conversions_produce_values() {
        unsafe {
            for (v, label) in [
                (VARIANT::i64(42), "BIGINT"),
                (VARIANT::i32(7), "INTEGER"),
                (VARIANT::f64(1.5), "DOUBLE"),
                (VARIANT::bool(true), "BOOLEAN"),
                (VARIANT::null(), "NULL"),
                (VARIANT::empty(), "空セル"),
                (VARIANT::date(45306.0), "DATE"),
            ] {
                let mut v = v;
                assert!(
                    variant_to_value(&v, label).is_ok(),
                    "{label} が変換できない"
                );
                v.clear();
            }
            let mut s = VARIANT::bstr("営業部");
            assert!(variant_to_value(&s, "VARCHAR").is_ok());
            s.clear();
        }
    }

    /// Excel の #N/A を黙って NULL にすると DB のデータが静かに壊れる。
    #[test]
    fn excel_error_cells_are_rejected_not_silently_nulled() {
        let mut v = VARIANT::empty();
        v.vt = VT_ERROR;
        v.value.lVal = -2146826246; // xlErrNA
        let err = unsafe { variant_to_value(&v, "2 行 3 列目") }.unwrap_err();
        assert!(err.contains("#N/A"), "{err}");
        assert!(err.contains("2 行 3 列目"), "位置が示されていない: {err}");
        assert!(err.contains("IFERROR"), "対処方法が示されていない: {err}");
    }

    #[test]
    fn objects_are_rejected_with_range_value_hint() {
        let mut v = VARIANT::empty();
        v.vt = VT_DISPATCH;
        let err = unsafe { variant_to_value(&v, "1 行 1 列目") }.unwrap_err();
        assert!(err.contains("Range.Value"), "{err}");
    }

    #[test]
    fn position_is_one_based_for_humans() {
        let (mut v, grid) = roundtrip(vec![vec![VARIANT::i32(1)]], 1);
        assert_eq!(grid.position(0, 0), "1 行 1 列目");
        assert_eq!(grid.position(2, 4), "3 行 5 列目");
        drop(grid);
        v.clear();
    }
}
