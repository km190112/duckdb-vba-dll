//! DuckDB のデータチャンク上の値 → VARIANT への変換。
//!
//! # 方針
//!
//! - スカラ型はすべてネイティブに対応する（Excel 利用者が扱うのはほぼこれ）。
//! - LIST / STRUCT / MAP などのコンテナ型は、セルに謎の文字列を入れるのではなく
//!   **列名と型名を明示してクエリ全体をエラーにする**。利用者は SQL 側で
//!   `::VARCHAR` にキャストすれば解決できる。黙って壊れたデータを貼るより良い。
//! - 数値の精度が f64 で失われる場合（HUGEINT、桁の大きい DECIMAL、u64 の上位域）は
//!   **黙って丸めず文字列にする**。金額を扱う用途で誤差が出るのが最悪なため。

#![allow(non_upper_case_globals)]

use crate::oleaut::*;
use crate::raw::Vector;
use libduckdb_sys as ffi;

/// OLE オートメーション日付の基準（1899-12-30）と Unix エポック（1970-01-01）の差（日）。
const OLE_EPOCH_OFFSET_DAYS: f64 = 25569.0;
const MICROS_PER_DAY: f64 = 86_400_000_000.0;

/// Excel が表示できる最小のシリアル値。これ未満は ISO 文字列にフォールバックする。
///
/// Excel は 1900-01-01 より前の日付を表示できず、OLE の負の日付は小数部の符号規則が
/// 特殊で事故が起きやすいので、そもそも数値にしない。
const MIN_DISPLAYABLE_OLE_DATE: f64 = 1.0;

/// この型は VARIANT に変換できないので、クエリごとエラーにする。
pub fn unsupported_type_name(t: ffi::duckdb_type) -> Option<&'static str> {
    use ffi::*;
    Some(match t {
        DUCKDB_TYPE_DUCKDB_TYPE_LIST => "LIST",
        DUCKDB_TYPE_DUCKDB_TYPE_STRUCT => "STRUCT",
        DUCKDB_TYPE_DUCKDB_TYPE_MAP => "MAP",
        DUCKDB_TYPE_DUCKDB_TYPE_ARRAY => "ARRAY",
        DUCKDB_TYPE_DUCKDB_TYPE_UNION => "UNION",
        DUCKDB_TYPE_DUCKDB_TYPE_GEOMETRY => "GEOMETRY",
        DUCKDB_TYPE_DUCKDB_TYPE_VARIANT => "VARIANT",
        DUCKDB_TYPE_DUCKDB_TYPE_BIT => "BIT",
        _ => return None,
    })
}

/// 変換できない列があれば、利用者が対処できるエラーメッセージを作る。
pub fn check_column_supported(col_name: &str, t: ffi::duckdb_type) -> Result<(), String> {
    if let Some(name) = unsupported_type_name(t) {
        return Err(format!(
            "列「{col_name}」の型 {name} は Excel のセルに変換できません。\n\
             SQL 側で文字列に変換してください（例: SELECT \"{col_name}\"::VARCHAR FROM ...）。"
        ));
    }
    Ok(())
}

/// ベクタの `row` 行目を VARIANT に変換する。
///
/// # Safety
/// `vec` は有効なベクタで、`row` はチャンクの行数未満であること。
pub unsafe fn cell_to_variant(vec: &Vector<'_>, row: usize) -> VARIANT {
    if !vec.is_valid(row) {
        return VARIANT::null();
    }

    use ffi::*;
    let t = vec.type_id();
    let data = vec.data_ptr();

    match t {
        DUCKDB_TYPE_DUCKDB_TYPE_BOOLEAN => VARIANT::bool(*(data as *const bool).add(row)),

        // i32 に収まる整数はすべて VT_I4。VBA の Long と一致する。
        DUCKDB_TYPE_DUCKDB_TYPE_TINYINT => VARIANT::i32(*(data as *const i8).add(row) as i32),
        DUCKDB_TYPE_DUCKDB_TYPE_SMALLINT => VARIANT::i32(*(data as *const i16).add(row) as i32),
        DUCKDB_TYPE_DUCKDB_TYPE_INTEGER => VARIANT::i32(*(data as *const i32).add(row)),
        DUCKDB_TYPE_DUCKDB_TYPE_UTINYINT => VARIANT::i32(*(data as *const u8).add(row) as i32),
        DUCKDB_TYPE_DUCKDB_TYPE_USMALLINT => VARIANT::i32(*(data as *const u16).add(row) as i32),
        DUCKDB_TYPE_DUCKDB_TYPE_UINTEGER => VARIANT::i64(*(data as *const u32).add(row) as i64),

        DUCKDB_TYPE_DUCKDB_TYPE_BIGINT => VARIANT::i64(*(data as *const i64).add(row)),
        DUCKDB_TYPE_DUCKDB_TYPE_UBIGINT => {
            let v = *(data as *const u64).add(row);
            // i64 に収まらない値を f64 に丸めると桁が落ちるので文字列にする
            if v <= i64::MAX as u64 {
                VARIANT::i64(v as i64)
            } else {
                VARIANT::bstr(&v.to_string())
            }
        }

        // HUGEINT は 128bit。f64 では表現しきれないので常に厳密な 10 進文字列にする。
        DUCKDB_TYPE_DUCKDB_TYPE_HUGEINT => {
            let h = *(data as *const duckdb_hugeint).add(row);
            VARIANT::bstr(&hugeint_to_string(h))
        }
        DUCKDB_TYPE_DUCKDB_TYPE_UHUGEINT => {
            let h = *(data as *const duckdb_uhugeint).add(row);
            let v = ((h.upper as u128) << 64) | (h.lower as u128);
            VARIANT::bstr(&v.to_string())
        }

        DUCKDB_TYPE_DUCKDB_TYPE_FLOAT => VARIANT::f64(*(data as *const f32).add(row) as f64),
        DUCKDB_TYPE_DUCKDB_TYPE_DOUBLE => VARIANT::f64(*(data as *const f64).add(row)),

        DUCKDB_TYPE_DUCKDB_TYPE_DECIMAL => decimal_to_variant(vec, data, row),

        DUCKDB_TYPE_DUCKDB_TYPE_VARCHAR | DUCKDB_TYPE_DUCKDB_TYPE_STRING_LITERAL => {
            let s = read_string_t(data as *mut duckdb_string_t, row);
            VARIANT::bstr(&String::from_utf8_lossy(&s))
        }

        DUCKDB_TYPE_DUCKDB_TYPE_BLOB => {
            let bytes = read_string_t(data as *mut duckdb_string_t, row);
            byte_array_variant(&bytes)
        }

        DUCKDB_TYPE_DUCKDB_TYPE_DATE => {
            let d = *(data as *const duckdb_date).add(row);
            ole_or_iso(d.days as f64 + OLE_EPOCH_OFFSET_DAYS, || iso_date(d))
        }

        DUCKDB_TYPE_DUCKDB_TYPE_TIME => {
            let t = *(data as *const duckdb_time).add(row);
            // 時刻のみ。OLE では小数部が時刻を表すので 0.0 起点でよい。
            VARIANT::date(t.micros as f64 / MICROS_PER_DAY)
        }
        DUCKDB_TYPE_DUCKDB_TYPE_TIME_TZ | DUCKDB_TYPE_DUCKDB_TYPE_TIME_NS => {
            // タイムゾーン付き時刻・ナノ秒時刻は情報が落ちるので文字列で返す
            VARIANT::bstr(&format!("{:?}", *(data as *const i64).add(row)))
        }

        DUCKDB_TYPE_DUCKDB_TYPE_TIMESTAMP
        | DUCKDB_TYPE_DUCKDB_TYPE_TIMESTAMP_TZ => {
            let ts = *(data as *const i64).add(row);
            timestamp_micros_to_variant(ts)
        }
        DUCKDB_TYPE_DUCKDB_TYPE_TIMESTAMP_S => {
            let s = *(data as *const i64).add(row);
            timestamp_micros_to_variant(s.saturating_mul(1_000_000))
        }
        DUCKDB_TYPE_DUCKDB_TYPE_TIMESTAMP_MS => {
            let ms = *(data as *const i64).add(row);
            timestamp_micros_to_variant(ms.saturating_mul(1_000))
        }
        DUCKDB_TYPE_DUCKDB_TYPE_TIMESTAMP_NS => {
            let ns = *(data as *const i64).add(row);
            timestamp_micros_to_variant(ns / 1_000)
        }

        DUCKDB_TYPE_DUCKDB_TYPE_UUID => {
            let h = *(data as *const duckdb_hugeint).add(row);
            VARIANT::bstr(&uuid_to_string(h))
        }

        DUCKDB_TYPE_DUCKDB_TYPE_INTERVAL => {
            let iv = *(data as *const duckdb_interval).add(row);
            VARIANT::bstr(&format!(
                "{} months {} days {} micros",
                iv.months, iv.days, iv.micros
            ))
        }

        DUCKDB_TYPE_DUCKDB_TYPE_ENUM => enum_to_variant(vec, data, row),

        DUCKDB_TYPE_DUCKDB_TYPE_SQLNULL => VARIANT::null(),

        // check_column_supported で事前に弾いているので、ここに来るのは
        // DuckDB に新しい型が増えた場合のみ。NULL を返して壊さない。
        _ => VARIANT::null(),
    }
}

/// エポックからのマイクロ秒を VT_DATE か ISO 文字列にする。
unsafe fn timestamp_micros_to_variant(micros: i64) -> VARIANT {
    let ole = micros as f64 / MICROS_PER_DAY + OLE_EPOCH_OFFSET_DAYS;
    ole_or_iso(ole, || iso_timestamp(micros))
}

/// Excel が表示できる範囲なら VT_DATE、そうでなければ文字列にフォールバックする。
unsafe fn ole_or_iso(ole: f64, fallback: impl FnOnce() -> String) -> VARIANT {
    if ole.is_finite() && (MIN_DISPLAYABLE_OLE_DATE..=2_958_465.0).contains(&ole) {
        VARIANT::date(ole)
    } else {
        VARIANT::bstr(&fallback())
    }
}

fn iso_date(d: ffi::duckdb_date) -> String {
    let s = unsafe { ffi::duckdb_from_date(d) };
    format!("{:04}-{:02}-{:02}", s.year, s.month, s.day)
}

fn iso_timestamp(micros: i64) -> String {
    let s = unsafe { ffi::duckdb_from_timestamp(ffi::duckdb_timestamp { micros }) };
    format!(
        "{:04}-{:02}-{:02} {:02}:{:02}:{:02}",
        s.date.year, s.date.month, s.date.day, s.time.hour, s.time.min, s.time.sec
    )
}

/// 128bit 符号付き整数を厳密な 10 進文字列にする。
fn hugeint_to_string(h: ffi::duckdb_hugeint) -> String {
    let v = ((h.upper as i128) << 64) | (h.lower as i128);
    v.to_string()
}

fn uuid_to_string(h: ffi::duckdb_hugeint) -> String {
    // DuckDB は UUID の上位ビットの符号ビットを反転して格納している
    let upper = (h.upper as u64) ^ (1u64 << 63);
    let b = ((upper as u128) << 64) | (h.lower as u128);
    let bytes = b.to_be_bytes();
    let hex: String = bytes.iter().map(|x| format!("{x:02x}")).collect();
    format!(
        "{}-{}-{}-{}-{}",
        &hex[0..8],
        &hex[8..12],
        &hex[12..16],
        &hex[16..20],
        &hex[20..32]
    )
}

/// f64 を返す DECIMAL の宣言桁数の上限。これを超えると厳密な文字列にする。
///
/// # このしきい値の根拠
///
/// Excel のセルは内部的に f64 なので、どう頑張っても f64 を超える精度は
/// セルに数値として入らない。よって選択肢は「黙って丸める」か「文字列にする」の 2 つ。
///
/// - `DECIMAL(18,2)` は金額列の宣言としてごく一般的で、実際の値は
///   せいぜい数十億（10^10 程度）に収まる。ここを文字列で返すと Excel 上で
///   左寄せのテキストになり `SUM()` が効かなくなる。**Excel に取り込む目的が壊れる**。
/// - 一方 `DECIMAL(19..38)` は i64 すら超える領域で、明確に「大きな数」を
///   意図した宣言。ここで黙って丸めると金額がずれる。
///
/// よって 18 を境界にする。18 以下は数値、19 以上は厳密な文字列。
const DECIMAL_MAX_WIDTH_AS_F64: u8 = 18;

/// DECIMAL の変換。
///
/// 宣言桁数が [`DECIMAL_MAX_WIDTH_AS_F64`] を超える場合は、黙って丸めず
/// 厳密な 10 進文字列にする。
unsafe fn decimal_to_variant(
    vec: &Vector<'_>,
    data: *mut std::ffi::c_void,
    row: usize,
) -> VARIANT {
    use ffi::*;
    let (width, scale) = vec.decimal_width_scale();

    let unscaled: i128 = match vec.decimal_internal_type() {
        DUCKDB_TYPE_DUCKDB_TYPE_SMALLINT => *(data as *const i16).add(row) as i128,
        DUCKDB_TYPE_DUCKDB_TYPE_INTEGER => *(data as *const i32).add(row) as i128,
        DUCKDB_TYPE_DUCKDB_TYPE_BIGINT => *(data as *const i64).add(row) as i128,
        DUCKDB_TYPE_DUCKDB_TYPE_HUGEINT => {
            let h = *(data as *const duckdb_hugeint).add(row);
            ((h.upper as i128) << 64) | (h.lower as i128)
        }
        _ => return VARIANT::null(),
    };

    if width <= DECIMAL_MAX_WIDTH_AS_F64 {
        VARIANT::f64(unscaled as f64 / 10f64.powi(scale as i32))
    } else {
        VARIANT::bstr(&decimal_to_string(unscaled, scale))
    }
}

/// スケール付き整数を正確な 10 進文字列にする（丸め無し）。
fn decimal_to_string(unscaled: i128, scale: u8) -> String {
    if scale == 0 {
        return unscaled.to_string();
    }
    let neg = unscaled < 0;
    let abs = unscaled.unsigned_abs();
    let digits = abs.to_string();
    let scale = scale as usize;
    let (int_part, frac_part) = if digits.len() > scale {
        let split = digits.len() - scale;
        (digits[..split].to_string(), digits[split..].to_string())
    } else {
        ("0".to_string(), format!("{:0>width$}", digits, width = scale))
    };
    format!("{}{}.{}", if neg { "-" } else { "" }, int_part, frac_part)
}

/// ENUM は辞書引きして文字列にする。
unsafe fn enum_to_variant(
    vec: &Vector<'_>,
    data: *mut std::ffi::c_void,
    row: usize,
) -> VARIANT {
    use ffi::*;
    let lt = vec.logical_type_raw();
    let index: u64 = match duckdb_enum_internal_type(lt) {
        DUCKDB_TYPE_DUCKDB_TYPE_UTINYINT => *(data as *const u8).add(row) as u64,
        DUCKDB_TYPE_DUCKDB_TYPE_USMALLINT => *(data as *const u16).add(row) as u64,
        DUCKDB_TYPE_DUCKDB_TYPE_UINTEGER => *(data as *const u32).add(row) as u64,
        _ => {
            let mut lt_mut = lt;
            duckdb_destroy_logical_type(&mut lt_mut);
            return VARIANT::null();
        }
    };
    let p = duckdb_enum_dictionary_value(lt, index);
    let s = if p.is_null() {
        String::new()
    } else {
        let owned = std::ffi::CStr::from_ptr(p).to_string_lossy().into_owned();
        duckdb_free(p as *mut _);
        owned
    };
    let mut lt_mut = lt;
    duckdb_destroy_logical_type(&mut lt_mut);
    VARIANT::bstr(&s)
}

/// `duckdb_string_t`（インライン格納とポインタ格納の 2 形態がある）からバイト列を読む。
unsafe fn read_string_t(base: *mut ffi::duckdb_string_t, row: usize) -> Vec<u8> {
    let p = base.add(row);
    let len = ffi::duckdb_string_t_length(*p) as usize;
    let data = ffi::duckdb_string_t_data(p);
    if data.is_null() || len == 0 {
        return Vec::new();
    }
    std::slice::from_raw_parts(data as *const u8, len).to_vec()
}

/// BLOB を VBA のバイト配列（`VT_ARRAY | VT_UI1`）にする。
fn byte_array_variant(bytes: &[u8]) -> VARIANT {
    // 0 バイトの SAFEARRAY は作れないので、空 BLOB は空文字列で表す
    if bytes.is_empty() {
        return VARIANT::bstr("");
    }
    let bounds = [SAFEARRAYBOUND { cElements: bytes.len() as u32, lLbound: 0 }];
    let psa = unsafe { SafeArrayCreate(VT_UI1, 1, bounds.as_ptr()) };
    if psa.is_null() {
        return VARIANT::null();
    }
    let mut data: *mut std::ffi::c_void = std::ptr::null_mut();
    if unsafe { SafeArrayAccessData(psa, &mut data) } != S_OK || data.is_null() {
        unsafe { SafeArrayDestroy(psa) };
        return VARIANT::null();
    }
    unsafe {
        std::ptr::copy_nonoverlapping(bytes.as_ptr(), data as *mut u8, bytes.len());
        SafeArrayUnaccessData(psa);
    }
    let mut u = VARIANT::empty();
    u.vt = VT_ARRAY | VT_UI1;
    u.value.parray = psa;
    u
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn container_types_are_rejected_with_actionable_message() {
        let err = check_column_supported("明細", ffi::DUCKDB_TYPE_DUCKDB_TYPE_LIST).unwrap_err();
        assert!(err.contains("LIST"), "{err}");
        assert!(err.contains("::VARCHAR"), "対処方法が示されていない: {err}");
        assert!(err.contains("明細"), "列名が示されていない: {err}");
    }

    #[test]
    fn scalar_types_are_accepted() {
        for t in [
            ffi::DUCKDB_TYPE_DUCKDB_TYPE_INTEGER,
            ffi::DUCKDB_TYPE_DUCKDB_TYPE_VARCHAR,
            ffi::DUCKDB_TYPE_DUCKDB_TYPE_TIMESTAMP,
            ffi::DUCKDB_TYPE_DUCKDB_TYPE_DECIMAL,
            ffi::DUCKDB_TYPE_DUCKDB_TYPE_BLOB,
        ] {
            assert!(check_column_supported("x", t).is_ok(), "型 {t} が拒否された");
        }
    }

    /// 高精度 DECIMAL を f64 に丸めないこと。金額計算で誤差が出るのを防ぐ。
    #[test]
    fn decimal_string_conversion_is_exact() {
        assert_eq!(decimal_to_string(123456789, 2), "1234567.89");
        assert_eq!(decimal_to_string(-123456789, 2), "-1234567.89");
        assert_eq!(decimal_to_string(5, 3), "0.005");
        assert_eq!(decimal_to_string(-5, 3), "-0.005");
        assert_eq!(decimal_to_string(1000, 0), "1000");
        assert_eq!(
            decimal_to_string(123456789012345678901234567890i128, 10),
            "12345678901234567890.1234567890"
        );
    }

    /// 金額列でよくある DECIMAL(18,2) は数値で返す（Excel で SUM が効くように）。
    /// それを超える宣言は黙って丸めず文字列にする。
    #[test]
    fn decimal_threshold_keeps_common_money_columns_numeric() {
        assert!(18 <= DECIMAL_MAX_WIDTH_AS_F64, "DECIMAL(18,2) は数値であるべき");
        assert!(19 > DECIMAL_MAX_WIDTH_AS_F64, "DECIMAL(19,x) は文字列であるべき");
    }

    #[test]
    fn hugeint_string_is_exact_not_rounded() {
        // f64 なら丸められてしまう大きさ
        let h = ffi::duckdb_hugeint { lower: 0, upper: 1 };
        assert_eq!(hugeint_to_string(h), "18446744073709551616");
        let neg = ffi::duckdb_hugeint { lower: u64::MAX, upper: -1 };
        assert_eq!(hugeint_to_string(neg), "-1");
    }

    #[test]
    fn ole_date_epoch_offset_is_correct() {
        // 1970-01-01 は Excel のシリアル値 25569
        assert_eq!(OLE_EPOCH_OFFSET_DAYS, 25569.0);
        // 2000-01-01 は Unix エポックから 10957 日 → 36526
        assert_eq!(10957.0 + OLE_EPOCH_OFFSET_DAYS, 36526.0);
    }

    #[test]
    fn empty_blob_becomes_empty_string_not_null_array() {
        let mut v = byte_array_variant(&[]);
        assert_eq!(v.vt, VT_BSTR);
        v.clear();
    }

    #[test]
    fn blob_becomes_byte_array() {
        let mut v = byte_array_variant(&[0x01, 0xFF, 0x00, 0x7F]);
        assert_eq!(v.vt, VT_ARRAY | VT_UI1);
        unsafe {
            let psa = v.value.parray;
            let mut lb = 0i32;
            let mut ub = 0i32;
            SafeArrayGetLBound(psa, 1, &mut lb);
            SafeArrayGetUBound(psa, 1, &mut ub);
            assert_eq!((lb, ub), (0, 3));
            let mut data: *mut std::ffi::c_void = std::ptr::null_mut();
            SafeArrayAccessData(psa, &mut data);
            let slice = std::slice::from_raw_parts(data as *const u8, 4);
            assert_eq!(slice, &[0x01, 0xFF, 0x00, 0x7F]);
            SafeArrayUnaccessData(psa);
        }
        v.clear();
    }
}
