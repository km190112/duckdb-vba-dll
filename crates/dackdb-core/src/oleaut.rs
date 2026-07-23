//! OLE Automation（`oleaut32.dll`）の最小限の生バインディング。
//!
//! `windows` クレートを使わず手書きしている理由：
//!
//! 1. 必要なのは 10 数個の関数と 2 つの構造体だけで、`windows` クレートは過剰。
//! 2. `windows` クレートの `VARIANT` は `Drop` を実装している。本 DLL は **VBA が所有する
//!    メモリに VARIANT を書き込む**のが仕事なので、Drop が走ると二重解放になり得る。
//!    ここでは Drop を持たない純粋な `#[repr(C)]` POD が欲しい。
//! 3. VARIANT / SAFEARRAY の ABI は 1996 年から凍結されており、バージョン追従の必要がない。
//!
//! 対象は x64 のみ（`sizeof(VARIANT) == 24`）。32bit ビルドは想定しない。

#![allow(non_snake_case, non_camel_case_types)]

use std::os::raw::c_void;

// ---------------------------------------------------------------------------
// VARTYPE 定数
// ---------------------------------------------------------------------------

pub const VT_EMPTY: u16 = 0;
pub const VT_NULL: u16 = 1;
pub const VT_I2: u16 = 2;
pub const VT_I4: u16 = 3;
pub const VT_R4: u16 = 4;
pub const VT_R8: u16 = 5;
pub const VT_DATE: u16 = 7;
pub const VT_BSTR: u16 = 8;
pub const VT_BOOL: u16 = 11;
pub const VT_VARIANT: u16 = 12;
pub const VT_UI1: u16 = 17;
pub const VT_I8: u16 = 20;
pub const VT_ARRAY: u16 = 0x2000;

// VBA / Excel から**入力として**渡ってくる可能性のある型。
pub const VT_I1: u16 = 16;
pub const VT_UI2: u16 = 18;
pub const VT_UI4: u16 = 19;
pub const VT_UI8: u16 = 21;
pub const VT_INT: u16 = 22;
pub const VT_UINT: u16 = 23;
pub const VT_CY: u16 = 6;
pub const VT_DECIMAL: u16 = 14;
/// Excel のセルが `#N/A` `#DIV/0!` などのエラー値のとき、`Range.Value` はこれを返す。
pub const VT_ERROR: u16 = 10;
pub const VT_DISPATCH: u16 = 9;
pub const VT_UNKNOWN: u16 = 13;
/// 値そのものではなく値へのポインタが入っていることを示すフラグ。
pub const VT_BYREF: u16 = 0x4000;
/// `vt` から VT_ARRAY / VT_BYREF を除いた基本型を取り出すマスク。
pub const VT_TYPEMASK: u16 = 0x0FFF;

/// VBA の `True`。VARIANT_BOOL は 0/-1 であって 0/1 ではない。
pub const VARIANT_TRUE: i16 = -1;
pub const VARIANT_FALSE: i16 = 0;

pub const S_OK: i32 = 0;

// ---------------------------------------------------------------------------
// 構造体
// ---------------------------------------------------------------------------

/// SAFEARRAY は不透明ポインタとして扱い、必ず API 経由で操作する。
#[repr(C)]
pub struct SAFEARRAY {
    _private: [u8; 0],
}

#[repr(C)]
#[derive(Clone, Copy)]
pub struct SAFEARRAYBOUND {
    pub cElements: u32,
    pub lLbound: i32,
}

/// VARIANT の値部分。x64 では 16 バイト（`BRECORD` が 2 ポインタ分あるため）。
///
/// `_pad` により必ず 16 バイト・アライメント 8 になることを保証する。
/// 小さいメンバ（`lval` など）を書いても上位バイトが残るため、
/// 構築時は必ず全体をゼロ初期化すること。
#[repr(C)]
#[derive(Clone, Copy)]
pub union VARIANT_VALUE {
    pub llVal: i64,
    pub lVal: i32,
    pub iVal: i16,
    pub bVal: u8,
    pub boolVal: i16,
    pub fltVal: f32,
    pub dblVal: f64,
    pub date: f64,
    pub bstrVal: *mut u16,
    pub parray: *mut SAFEARRAY,
    pub byref: *mut c_void,
    _pad: [u8; 16],
}

/// `#[repr(C)]` の POD。**`Drop` を実装しないこと**（VBA 所有のメモリに書き込むため）。
#[repr(C)]
#[derive(Clone, Copy)]
pub struct VARIANT {
    pub vt: u16,
    pub wReserved1: u16,
    pub wReserved2: u16,
    pub wReserved3: u16,
    pub value: VARIANT_VALUE,
}

impl VARIANT {
    /// VT_EMPTY のゼロ初期化された VARIANT。
    pub fn empty() -> Self {
        // union 全体をゼロで埋める。中途半端なメンバ書き込みによるゴミを防ぐ。
        VARIANT {
            vt: VT_EMPTY,
            wReserved1: 0,
            wReserved2: 0,
            wReserved3: 0,
            value: VARIANT_VALUE { _pad: [0u8; 16] },
        }
    }

    fn of(vt: u16, value: VARIANT_VALUE) -> Self {
        VARIANT { vt, wReserved1: 0, wReserved2: 0, wReserved3: 0, value }
    }

    /// SQL の NULL。VBA の `IsNull()` が True になり、セルは空欄になる。
    pub fn null() -> Self {
        Self::of(VT_NULL, VARIANT_VALUE { _pad: [0u8; 16] })
    }

    pub fn i32(v: i32) -> Self {
        let mut u = VARIANT_VALUE { _pad: [0u8; 16] };
        u.lVal = v;
        Self::of(VT_I4, u)
    }

    pub fn i64(v: i64) -> Self {
        let mut u = VARIANT_VALUE { _pad: [0u8; 16] };
        u.llVal = v;
        Self::of(VT_I8, u)
    }

    pub fn f64(v: f64) -> Self {
        let mut u = VARIANT_VALUE { _pad: [0u8; 16] };
        u.dblVal = v;
        Self::of(VT_R8, u)
    }

    /// OLE Automation 日付（1899-12-30 を 0 とする日数）。
    pub fn date(v: f64) -> Self {
        let mut u = VARIANT_VALUE { _pad: [0u8; 16] };
        u.date = v;
        Self::of(VT_DATE, u)
    }

    pub fn bool(v: bool) -> Self {
        let mut u = VARIANT_VALUE { _pad: [0u8; 16] };
        u.boolVal = if v { VARIANT_TRUE } else { VARIANT_FALSE };
        Self::of(VT_BOOL, u)
    }

    /// Rust の `&str` から BSTR を確保した VARIANT を作る。
    ///
    /// UTF-8 → UTF-16 変換はここで行う。**日本語対応の要**であり、
    /// VBA の `Declare` による ANSI マーシャリングを一切経由しない。
    pub fn bstr(s: &str) -> Self {
        let wide: Vec<u16> = s.encode_utf16().collect();
        let b = unsafe { SysAllocStringLen(wide.as_ptr(), wide.len() as u32) };
        let mut u = VARIANT_VALUE { _pad: [0u8; 16] };
        u.bstrVal = b;
        // 確保に失敗した場合は VT_NULL に落とす（null BSTR を VBA に渡さない）
        if b.is_null() {
            return Self::null();
        }
        Self::of(VT_BSTR, u)
    }

    /// SAFEARRAY の所有権を VARIANT に移す。
    pub fn array(psa: *mut SAFEARRAY) -> Self {
        let mut u = VARIANT_VALUE { _pad: [0u8; 16] };
        u.parray = psa;
        Self::of(VT_ARRAY | VT_VARIANT, u)
    }

    /// この VARIANT が保持しているリソース（BSTR / SAFEARRAY）を解放する。
    ///
    /// VBA に引き渡した VARIANT に対しては**呼んではいけない**（所有権は VBA に移っている）。
    /// 引き渡しに失敗した場合の後始末や一時 VARIANT の破棄にのみ使う。
    pub fn clear(&mut self) {
        unsafe {
            VariantClear(self);
        }
    }
}

/// 診断用。`vt` に応じて中身を安全に読む（union の誤読を避けるため必ず vt で分岐する）。
impl std::fmt::Debug for VARIANT {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        unsafe {
            match self.vt {
                VT_EMPTY => write!(f, "VARIANT(EMPTY)"),
                VT_NULL => write!(f, "VARIANT(NULL)"),
                VT_I4 => write!(f, "VARIANT(I4 {})", self.value.lVal),
                VT_I8 => write!(f, "VARIANT(I8 {})", self.value.llVal),
                VT_R8 => write!(f, "VARIANT(R8 {})", self.value.dblVal),
                VT_DATE => write!(f, "VARIANT(DATE {})", self.value.date),
                VT_BOOL => write!(f, "VARIANT(BOOL {})", self.value.boolVal != 0),
                VT_BSTR => write!(f, "VARIANT(BSTR {:?})", bstr_to_string(self.value.bstrVal)),
                v if v & VT_ARRAY != 0 => write!(f, "VARIANT(ARRAY vt=0x{v:04X})"),
                v => write!(f, "VARIANT(vt=0x{v:04X})"),
            }
        }
    }
}

// ---------------------------------------------------------------------------
// oleaut32 のエクスポート
// ---------------------------------------------------------------------------

#[link(name = "oleaut32")]
extern "system" {
    pub fn SysAllocStringLen(psz: *const u16, len: u32) -> *mut u16;
    pub fn SysFreeString(bstr: *mut u16);
    pub fn SysStringLen(bstr: *mut u16) -> u32;

    pub fn VariantInit(pvarg: *mut VARIANT);
    pub fn VariantClear(pvarg: *mut VARIANT) -> i32;
    pub fn VariantCopy(dest: *mut VARIANT, src: *const VARIANT) -> i32;
    pub fn VariantChangeType(dest: *mut VARIANT, src: *const VARIANT, flags: u16, vt: u16) -> i32;

    pub fn SafeArrayCreate(vt: u16, cDims: u32, rgsabound: *const SAFEARRAYBOUND)
        -> *mut SAFEARRAY;
    pub fn SafeArrayDestroy(psa: *mut SAFEARRAY) -> i32;
    pub fn SafeArrayGetDim(psa: *mut SAFEARRAY) -> u32;
    pub fn SafeArrayGetLBound(psa: *mut SAFEARRAY, nDim: u32, plLbound: *mut i32) -> i32;
    pub fn SafeArrayGetUBound(psa: *mut SAFEARRAY, nDim: u32, plUbound: *mut i32) -> i32;
    pub fn SafeArrayGetVartype(psa: *mut SAFEARRAY, pvt: *mut u16) -> i32;
    pub fn SafeArrayGetElement(psa: *mut SAFEARRAY, rgIndices: *const i32, pv: *mut c_void)
        -> i32;
    pub fn SafeArrayPutElement(psa: *mut SAFEARRAY, rgIndices: *const i32, pv: *const c_void)
        -> i32;
    pub fn SafeArrayAccessData(psa: *mut SAFEARRAY, ppvData: *mut *mut c_void) -> i32;
    pub fn SafeArrayUnaccessData(psa: *mut SAFEARRAY) -> i32;
}

/// BSTR を Rust の `String` に変換する（所有権は移動しない）。
///
/// # Safety
/// `bstr` は有効な BSTR か null であること。
pub unsafe fn bstr_to_string(bstr: *mut u16) -> String {
    if bstr.is_null() {
        return String::new();
    }
    let len = SysStringLen(bstr) as usize;
    let slice = std::slice::from_raw_parts(bstr, len);
    String::from_utf16_lossy(slice)
}

/// VBA が `StrPtr()` で渡してくる NUL 終端の UTF-16 文字列を `String` にする。
///
/// # Safety
/// `p` は NUL 終端の有効な UTF-16 文字列か null であること。
pub unsafe fn wide_ptr_to_string(p: *const u16) -> Option<String> {
    if p.is_null() {
        return None;
    }
    let mut len = 0usize;
    // 暴走防止の上限（SQL 文が 64MiB を超えることはない）
    const MAX: usize = 32 * 1024 * 1024;
    while len < MAX && *p.add(len) != 0 {
        len += 1;
    }
    let slice = std::slice::from_raw_parts(p, len);
    Some(String::from_utf16_lossy(slice))
}

#[cfg(test)]
mod tests {
    use super::*;

    /// x64 における VARIANT の ABI サイズ。ここがズレると VBA 側のスタックが壊れる。
    #[test]
    fn variant_abi_layout_is_24_bytes_on_x64() {
        assert_eq!(std::mem::size_of::<VARIANT>(), 24);
        assert_eq!(std::mem::align_of::<VARIANT>(), 8);
        assert_eq!(std::mem::size_of::<VARIANT_VALUE>(), 16);
        // 値部分は必ずオフセット 8 から始まる
        let v = VARIANT::empty();
        let base = &v as *const VARIANT as usize;
        let val = &v.value as *const VARIANT_VALUE as usize;
        assert_eq!(val - base, 8);
    }

    #[test]
    fn safearraybound_layout() {
        assert_eq!(std::mem::size_of::<SAFEARRAYBOUND>(), 8);
    }

    #[test]
    fn bstr_roundtrip_preserves_japanese_and_astral_chars() {
        // 漢字・ひらがな・サロゲートペア（𠮷）・絵文字
        let src = "テスト漢字𠮷😀";
        let mut v = VARIANT::bstr(src);
        assert_eq!(v.vt, VT_BSTR);
        let back = unsafe { bstr_to_string(v.value.bstrVal) };
        assert_eq!(back, src);
        v.clear();
    }

    #[test]
    fn empty_string_is_valid_bstr_not_null() {
        let mut v = VARIANT::bstr("");
        assert_eq!(v.vt, VT_BSTR);
        assert!(unsafe { !v.value.bstrVal.is_null() });
        assert_eq!(unsafe { bstr_to_string(v.value.bstrVal) }, "");
        v.clear();
    }

    #[test]
    fn scalar_constructors_zero_the_unused_bytes() {
        // i32 を入れた後の上位 4 バイトにゴミが残っていないこと
        let v = VARIANT::i32(-1);
        assert_eq!(v.vt, VT_I4);
        assert_eq!(unsafe { v.value.lVal }, -1);
        assert_eq!(unsafe { v.value.llVal }, 0xFFFF_FFFFi64); // 上位はゼロのまま

        let b = VARIANT::bool(true);
        assert_eq!(unsafe { b.value.boolVal }, VARIANT_TRUE);
        assert_eq!(unsafe { b.value.llVal }, 0xFFFFi64);
    }

    #[test]
    fn wide_ptr_roundtrip() {
        let s = "SELECT * FROM 売上";
        let mut wide: Vec<u16> = s.encode_utf16().collect();
        wide.push(0);
        let got = unsafe { wide_ptr_to_string(wide.as_ptr()) };
        assert_eq!(got.as_deref(), Some(s));
        assert_eq!(unsafe { wide_ptr_to_string(std::ptr::null()) }, None);
    }
}
