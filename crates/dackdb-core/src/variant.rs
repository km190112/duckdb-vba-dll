//! 2次元 SAFEARRAY（VBA の `Variant` 2次元配列）の組み立て。
//!
//! VBA 側は `Range("A1").Resize(r, c).Value = arr` の 1 行でシートに貼れる。

use crate::oleaut::*;

/// Excel シートの上限。これを超える結果セットは貼り付けられないので拒否する。
pub const MAX_EXCEL_ROWS: usize = 1_048_576;
pub const MAX_EXCEL_COLS: usize = 16_384;

/// 行優先で VARIANT を溜めていく組み立て用バッファ。
///
/// SAFEARRAY は列優先（Fortran 順）で格納されるため、`into_safearray` で並べ替える。
pub struct Grid {
    cols: usize,
    /// 行優先。長さは必ず `rows * cols`。
    cells: Vec<VARIANT>,
}

impl Grid {
    /// ヘッダ行（列名）から開始する。以降 `push_row` で 1 行ずつ追加する。
    pub fn with_header(headers: &[String]) -> Self {
        let cols = headers.len();
        let mut cells = Vec::with_capacity(cols);
        for h in headers {
            cells.push(VARIANT::bstr(h));
        }
        Grid { cols, cells }
    }

    /// 列数だけ決めてヘッダ無しで開始する。
    pub fn with_cols(cols: usize) -> Self {
        Grid { cols, cells: Vec::new() }
    }

    pub fn cols(&self) -> usize {
        self.cols
    }

    pub fn rows(&self) -> usize {
        if self.cols == 0 {
            0
        } else {
            self.cells.len() / self.cols
        }
    }

    /// 1 セル追加する。呼び出し側が行あたり `cols` 個ちょうど積むこと。
    pub fn push(&mut self, v: VARIANT) {
        self.cells.push(v);
    }

    /// 行を丸ごと追加する。
    ///
    /// # Panics
    /// `row.len() != self.cols` のとき。
    pub fn push_row(&mut self, row: Vec<VARIANT>) {
        assert_eq!(row.len(), self.cols, "行の列数が Grid の列数と一致しない");
        self.cells.extend(row);
    }

    /// Excel の上限を超えていないか検査する。
    pub fn check_excel_limits(&self) -> Result<(), String> {
        let rows = self.rows();
        if rows > MAX_EXCEL_ROWS {
            return Err(format!(
                "結果が {rows} 行あり、Excel シートの上限 {MAX_EXCEL_ROWS} 行を超えています。\
                 SQL に LIMIT を付けて絞り込んでください。"
            ));
        }
        if self.cols > MAX_EXCEL_COLS {
            return Err(format!(
                "結果が {} 列あり、Excel シートの上限 {MAX_EXCEL_COLS} 列を超えています。",
                self.cols
            ));
        }
        Ok(())
    }

    /// 下限 1 の 2 次元 SAFEARRAY(VARIANT) に変換し、所有権ごと VARIANT に包んで返す。
    ///
    /// 成功時、保持していた BSTR / 入れ子配列の所有権はすべて SAFEARRAY に移る。
    /// 失敗時は自身の VARIANT をすべて解放してから `Err` を返すのでリークしない。
    pub fn into_variant(mut self) -> Result<VARIANT, String> {
        self.check_excel_limits()?;

        let rows = self.rows();
        let cols = self.cols;

        // 0 行 0 列は SAFEARRAY として作れないため、呼び出し側で空判定を済ませておく。
        if rows == 0 || cols == 0 {
            self.clear_all();
            return Err("結果が空です（0 行または 0 列）".to_string());
        }

        // rgsabound[0] が最も左（＝最も速く変化する）次元。
        // VBA の arr(row, col) に合わせて 次元1 = 行、次元2 = 列 とする。
        let bounds = [
            SAFEARRAYBOUND { cElements: rows as u32, lLbound: 1 },
            SAFEARRAYBOUND { cElements: cols as u32, lLbound: 1 },
        ];

        let psa = unsafe { SafeArrayCreate(VT_VARIANT, 2, bounds.as_ptr()) };
        if psa.is_null() {
            self.clear_all();
            return Err(format!(
                "SAFEARRAY の確保に失敗しました（{rows} 行 x {cols} 列）。メモリ不足の可能性があります。"
            ));
        }

        let mut data: *mut std::ffi::c_void = std::ptr::null_mut();
        let hr = unsafe { SafeArrayAccessData(psa, &mut data) };
        if hr != S_OK || data.is_null() {
            unsafe { SafeArrayDestroy(psa) };
            self.clear_all();
            return Err(format!("SafeArrayAccessData に失敗しました (HRESULT=0x{hr:08X})"));
        }

        // SAFEARRAY は列優先。要素 (r, c) の線形位置は c * rows + r。
        // SafeArrayCreate がデータをゼロ初期化しているので、上書きは move であって
        // 既存要素の解放漏れにはならない。
        let dst = data as *mut VARIANT;
        for (idx, v) in self.cells.drain(..).enumerate() {
            let r = idx / cols;
            let c = idx % cols;
            unsafe { dst.add(c * rows + r).write(v) };
        }

        unsafe { SafeArrayUnaccessData(psa) };
        Ok(VARIANT::array(psa))
    }

    /// 溜め込んだ VARIANT をすべて解放する（失敗パス用）。
    fn clear_all(&mut self) {
        for v in self.cells.iter_mut() {
            v.clear();
        }
        self.cells.clear();
    }
}

impl Drop for Grid {
    /// `into_variant` を通らずに破棄された場合に BSTR をリークさせない。
    fn drop(&mut self) {
        self.clear_all();
    }
}

/// 単一の値を VBA へ返すときのヘルパ。1x1 の配列ではなくスカラ VARIANT を返す。
pub fn scalar_i64(v: i64) -> VARIANT {
    VARIANT::i64(v)
}

/// 呼び出し側から渡された出力用 VARIANT に値を書き込む。
///
/// VBA は初期化済みの Variant を渡してくるため、**上書き前に必ず `VariantClear` する**。
/// これを忘れると VBA 側で以前入っていた文字列や配列がリークする。
///
/// # Safety
/// `out` は VBA が所有する有効な `VARIANT*` であること。
pub unsafe fn write_out(out: *mut VARIANT, value: VARIANT) {
    if out.is_null() {
        // 受け取り手がいないので value 自身を解放して捨てる
        let mut v = value;
        v.clear();
        return;
    }
    VariantClear(out);
    *out = value;
}

#[cfg(test)]
mod tests {
    use super::*;

    /// 指定位置の要素を SafeArrayGetElement で読み直す（API 経由の独立検証）。
    unsafe fn get(psa: *mut SAFEARRAY, row: i32, col: i32) -> VARIANT {
        let mut out = VARIANT::empty();
        VariantInit(&mut out);
        let idx = [row, col];
        let hr = SafeArrayGetElement(psa, idx.as_ptr(), &mut out as *mut VARIANT as *mut _);
        assert_eq!(hr, S_OK, "SafeArrayGetElement が失敗 ({row}, {col})");
        out
    }

    /// 手書きの列優先インデックス計算が正しいことを、OLE の公式 API と突き合わせて証明する。
    /// SAFEARRAY は C の配列と違い列優先なので、ここを間違えると行と列が入れ替わる。
    #[test]
    fn column_major_indexing_matches_safearray_api() {
        let mut g = Grid::with_header(&["A".into(), "B".into(), "C".into()]);
        g.push_row(vec![VARIANT::i32(10), VARIANT::i32(11), VARIANT::i32(12)]);
        g.push_row(vec![VARIANT::i32(20), VARIANT::i32(21), VARIANT::i32(22)]);
        assert_eq!(g.rows(), 3);
        assert_eq!(g.cols(), 3);

        let mut v = g.into_variant().expect("変換に失敗");
        let psa = unsafe { v.value.parray };

        unsafe {
            // 次元1 = 行（1..3）、次元2 = 列（1..3）
            let mut lb = 0i32;
            let mut ub = 0i32;
            SafeArrayGetLBound(psa, 1, &mut lb);
            SafeArrayGetUBound(psa, 1, &mut ub);
            assert_eq!((lb, ub), (1, 3), "次元1（行）の範囲");
            SafeArrayGetLBound(psa, 2, &mut lb);
            SafeArrayGetUBound(psa, 2, &mut ub);
            assert_eq!((lb, ub), (1, 3), "次元2（列）の範囲");

            // ヘッダ行
            let mut h = get(psa, 1, 2);
            assert_eq!(h.vt, VT_BSTR);
            assert_eq!(bstr_to_string(h.value.bstrVal), "B");
            h.clear();

            // データ行：(2,1)=10, (2,3)=12, (3,1)=20, (3,3)=22
            assert_eq!(get(psa, 2, 1).value.lVal, 10);
            assert_eq!(get(psa, 2, 3).value.lVal, 12);
            assert_eq!(get(psa, 3, 1).value.lVal, 20);
            assert_eq!(get(psa, 3, 3).value.lVal, 22);
        }

        v.clear();
    }

    #[test]
    fn japanese_and_null_survive_the_grid() {
        let mut g = Grid::with_header(&["名前".into(), "備考".into()]);
        g.push_row(vec![VARIANT::bstr("テスト漢字𠮷😀"), VARIANT::null()]);

        let mut v = g.into_variant().expect("変換に失敗");
        unsafe {
            let psa = v.value.parray;
            let mut name = get(psa, 2, 1);
            assert_eq!(bstr_to_string(name.value.bstrVal), "テスト漢字𠮷😀");
            name.clear();

            let mut header = get(psa, 1, 1);
            assert_eq!(bstr_to_string(header.value.bstrVal), "名前");
            header.clear();

            // NULL は VT_NULL のまま（VBA の IsNull() が True になる）
            assert_eq!(get(psa, 2, 2).vt, VT_NULL);
        }
        v.clear();
    }

    #[test]
    fn single_cell_grid_works() {
        let mut g = Grid::with_cols(1);
        g.push_row(vec![VARIANT::i64(42)]);
        let mut v = g.into_variant().expect("変換に失敗");
        unsafe {
            assert_eq!(get(v.value.parray, 1, 1).value.llVal, 42);
        }
        v.clear();
    }

    #[test]
    fn empty_grid_is_rejected_not_crashed() {
        let g = Grid::with_cols(0);
        assert!(g.into_variant().is_err());
    }

    #[test]
    fn too_many_columns_is_rejected() {
        let g = Grid::with_cols(MAX_EXCEL_COLS + 1);
        let err = g.check_excel_limits().unwrap_err();
        assert!(err.contains("列"), "エラーメッセージ: {err}");
    }

    /// `into_variant` を呼ばずに Grid を捨てても BSTR がリークしないこと。
    /// （リーク自体はテストで直接観測できないので、Drop が走ることだけ確認する）
    #[test]
    fn dropping_grid_without_conversion_does_not_panic() {
        let mut g = Grid::with_header(&["あ".into(), "い".into()]);
        g.push_row(vec![VARIANT::bstr("値1"), VARIANT::bstr("値2")]);
        drop(g);
    }

    #[test]
    fn write_out_clears_previous_value_before_overwriting() {
        // VBA が使い回した Variant を模す：先に文字列を入れておく
        let mut slot = VARIANT::bstr("以前の値");
        assert_eq!(slot.vt, VT_BSTR);

        unsafe { write_out(&mut slot, VARIANT::i64(7)) };
        assert_eq!(slot.vt, VT_I8);
        assert_eq!(unsafe { slot.value.llVal }, 7);
    }

    #[test]
    fn write_out_to_null_pointer_frees_the_value() {
        // 受け取り手が null でもリークせず、クラッシュもしない
        unsafe { write_out(std::ptr::null_mut(), VARIANT::bstr("捨てられる値")) };
    }
}
