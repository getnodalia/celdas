//! VBA values (`Variant`), their conversions and the run-time error type.

use std::any::Any;
use std::cell::RefCell;
use std::rc::Rc;

/// A VBA run-time error: its number and description, as `Err.Number` and `Err.Description`.
#[derive(Clone, Debug, PartialEq)]
pub struct VbError {
    pub number: i32,
    pub description: String,
}

pub type Result<T> = std::result::Result<T, VbError>;

impl VbError {
    pub fn new(number: i32, description: impl Into<String>) -> VbError {
        VbError { number, description: description.into() }
    }
    pub fn type_mismatch() -> VbError {
        VbError::new(13, "Type mismatch")
    }
    pub fn invalid_call() -> VbError {
        VbError::new(5, "Invalid procedure call or argument")
    }
    pub fn overflow() -> VbError {
        VbError::new(6, "Overflow")
    }
    pub fn div_zero() -> VbError {
        VbError::new(11, "Division by zero")
    }
    pub fn subscript() -> VbError {
        VbError::new(9, "Subscript out of range")
    }
    pub fn object_required() -> VbError {
        VbError::new(424, "Object required")
    }
    pub fn not_set() -> VbError {
        VbError::new(91, "Object variable or With block variable not set")
    }
    pub fn no_member(name: &str) -> VbError {
        VbError::new(438, format!("Object doesn't support this property or method: '{name}'"))
    }
    pub fn not_defined(name: &str) -> VbError {
        VbError::new(35, format!("Sub or Function not defined: '{name}'"))
    }
    /// An error of the object model (Excel reports 1004 for most of them).
    pub fn app(description: impl Into<String>) -> VbError {
        VbError::new(1004, description)
    }
}

impl std::fmt::Display for VbError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(f, "Run-time error '{}': {}", self.number, self.description)
    }
}

impl std::error::Error for VbError {}

/// An object a host provides (a worksheet, a range, a chart…). The interpreter keeps it opaque
/// and hands it back to the host for every member access.
pub trait HostObject: Any {
    /// The class name `TypeName` reports (`Range`, `Worksheet`…).
    fn class(&self) -> String;
    fn as_any(&self) -> &dyn Any;
    /// Whether `other` is the same object (`Is`). Two handles to one worksheet are the same.
    fn same(&self, other: &dyn HostObject) -> bool;
}

/// A shared object value.
#[derive(Clone)]
pub struct Obj(pub Rc<dyn HostObject>);

impl Obj {
    pub fn new(o: impl HostObject) -> Obj {
        Obj(Rc::new(o))
    }
    pub fn class(&self) -> String {
        self.0.class()
    }
    pub fn downcast<T: 'static>(&self) -> Option<&T> {
        self.0.as_any().downcast_ref::<T>()
    }
    pub fn same(&self, other: &Obj) -> bool {
        Rc::ptr_eq(&self.0, &other.0) || self.0.same(other.0.as_ref())
    }
}

impl std::fmt::Debug for Obj {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(f, "<{}>", self.class())
    }
}

/// A VBA array: bounds per dimension and its elements in row-major order.
#[derive(Clone, Debug, PartialEq)]
pub struct VArray {
    pub bounds: Vec<(i32, i32)>,
    pub data: Vec<Value>,
}

/// Most elements an array may hold (a hostile `ReDim` must not exhaust memory).
pub const MAX_ARRAY_ELEMENTS: usize = 50_000_000;

impl VArray {
    pub fn new(bounds: Vec<(i32, i32)>, fill: Value) -> Result<VArray> {
        let mut total: usize = 1;
        for &(lo, hi) in &bounds {
            if hi < lo {
                return Err(VbError::subscript());
            }
            let n = (i64::from(hi) - i64::from(lo) + 1) as usize;
            total = total.checked_mul(n).filter(|t| *t <= MAX_ARRAY_ELEMENTS).ok_or_else(|| VbError::new(7, "Out of memory"))?;
        }
        Ok(VArray { bounds, data: vec![fill; total] })
    }

    /// A one-dimensional array holding `items`, from `lo`.
    pub fn from_vec(items: Vec<Value>, lo: i32) -> VArray {
        let hi = lo.saturating_add(items.len() as i32).saturating_sub(1);
        VArray { bounds: vec![(lo, hi)], data: items }
    }

    pub fn offset(&self, idx: &[i32]) -> Result<usize> {
        if idx.len() != self.bounds.len() {
            return Err(VbError::subscript());
        }
        let mut off: usize = 0;
        for (&i, &(lo, hi)) in idx.iter().zip(&self.bounds) {
            if i < lo || i > hi {
                return Err(VbError::subscript());
            }
            let n = (i64::from(hi) - i64::from(lo) + 1) as usize;
            off = off
                .checked_mul(n)
                .ok_or_else(VbError::subscript)?
                .checked_add((i64::from(i) - i64::from(lo)) as usize)
                .ok_or_else(VbError::subscript)?;
        }
        Ok(off)
    }

    pub fn get(&self, idx: &[i32]) -> Result<Value> {
        let off = self.offset(idx)?;
        self.data.get(off).cloned().ok_or_else(VbError::subscript)
    }

    pub fn set(&mut self, idx: &[i32], v: Value) -> Result<()> {
        let off = self.offset(idx)?;
        match self.data.get_mut(off) {
            Some(slot) => {
                *slot = v;
                Ok(())
            }
            None => Err(VbError::subscript()),
        }
    }

    pub fn len(&self) -> usize {
        self.data.len()
    }

    pub fn is_empty(&self) -> bool {
        self.data.is_empty()
    }
}

/// A VBA `Variant`.
#[derive(Clone, Debug)]
pub enum Value {
    Empty,
    Null,
    Bool(bool),
    Byte(u8),
    Int(i16),
    Long(i32),
    Single(f32),
    Double(f64),
    /// A date and time as an Excel serial number (1900 system).
    Date(f64),
    Str(Rc<str>),
    /// A cell error value (`CVErr`), by its number (2007 is `#DIV/0!`).
    Error(i32),
    Array(Rc<RefCell<VArray>>),
    /// An object, or `Nothing`.
    Object(Option<Obj>),
    /// An omitted optional argument.
    Missing,
}

impl PartialEq for Value {
    fn eq(&self, other: &Value) -> bool {
        match (self, other) {
            (Value::Empty, Value::Empty) | (Value::Null, Value::Null) | (Value::Missing, Value::Missing) => true,
            (Value::Str(a), Value::Str(b)) => a == b,
            (Value::Bool(a), Value::Bool(b)) => a == b,
            (Value::Error(a), Value::Error(b)) => a == b,
            (Value::Object(a), Value::Object(b)) => match (a, b) {
                (None, None) => true,
                (Some(a), Some(b)) => a.same(b),
                _ => false,
            },
            (Value::Array(a), Value::Array(b)) => Rc::ptr_eq(a, b),
            (a, b) => match (a.as_number(), b.as_number()) {
                (Some(x), Some(y)) => x == y,
                _ => false,
            },
        }
    }
}

impl Value {
    pub fn str(s: &str) -> Value {
        Value::Str(Rc::from(s))
    }
    pub fn string(s: String) -> Value {
        Value::Str(Rc::from(s))
    }
    pub fn nothing() -> Value {
        Value::Object(None)
    }
    pub fn object(o: impl HostObject) -> Value {
        Value::Object(Some(Obj::new(o)))
    }
    pub fn array(a: VArray) -> Value {
        Value::Array(Rc::new(RefCell::new(a)))
    }

    /// The numeric content of a number, boolean or date (not of a string).
    pub fn as_number(&self) -> Option<f64> {
        match self {
            Value::Bool(b) => Some(if *b { -1.0 } else { 0.0 }),
            Value::Byte(n) => Some(f64::from(*n)),
            Value::Int(n) => Some(f64::from(*n)),
            Value::Long(n) => Some(f64::from(*n)),
            Value::Single(n) => Some(f64::from(*n)),
            Value::Double(n) | Value::Date(n) => Some(*n),
            _ => None,
        }
    }

    pub fn is_numeric_type(&self) -> bool {
        matches!(self, Value::Byte(_) | Value::Int(_) | Value::Long(_) | Value::Single(_) | Value::Double(_))
    }

    pub fn is_empty_like(&self) -> bool {
        matches!(self, Value::Empty | Value::Missing)
    }

    /// `VarType`.
    pub fn var_type(&self) -> i32 {
        match self {
            Value::Empty | Value::Missing => 0,
            Value::Null => 1,
            Value::Int(_) => 2,
            Value::Long(_) => 3,
            Value::Single(_) => 4,
            Value::Double(_) => 5,
            Value::Date(_) => 7,
            Value::Str(_) => 8,
            Value::Object(_) => 9,
            Value::Error(_) => 10,
            Value::Bool(_) => 11,
            Value::Byte(_) => 17,
            Value::Array(_) => 8192 | 12,
        }
    }

    /// `TypeName`.
    pub fn type_name(&self) -> String {
        match self {
            Value::Empty | Value::Missing => "Empty".into(),
            Value::Null => "Null".into(),
            Value::Bool(_) => "Boolean".into(),
            Value::Byte(_) => "Byte".into(),
            Value::Int(_) => "Integer".into(),
            Value::Long(_) => "Long".into(),
            Value::Single(_) => "Single".into(),
            Value::Double(_) => "Double".into(),
            Value::Date(_) => "Date".into(),
            Value::Str(_) => "String".into(),
            Value::Error(_) => "Error".into(),
            Value::Array(_) => "Variant()".into(),
            Value::Object(None) => "Nothing".into(),
            Value::Object(Some(o)) => o.class(),
        }
    }

    /// The value as a number (`CDbl`): strings are read as numbers, `Empty` is 0.
    pub fn to_f64(&self) -> Result<f64> {
        match self {
            Value::Empty | Value::Missing => Ok(0.0),
            Value::Str(s) => parse_number(s).ok_or_else(VbError::type_mismatch),
            Value::Error(_) | Value::Null | Value::Array(_) | Value::Object(_) => Err(VbError::type_mismatch()),
            other => other.as_number().ok_or_else(VbError::type_mismatch),
        }
    }

    /// The value as a `Long` (`CLng`): rounded half to even, with an overflow error.
    pub fn to_i32(&self) -> Result<i32> {
        let n = round_even(self.to_f64()?);
        if n.is_nan() || !(-2_147_483_648.0..=2_147_483_647.0).contains(&n) {
            return Err(VbError::overflow());
        }
        Ok(n as i32)
    }

    pub fn to_i64(&self) -> Result<i64> {
        let n = round_even(self.to_f64()?);
        if n.is_nan() || n.abs() > 9.0e15 {
            return Err(VbError::overflow());
        }
        Ok(n as i64)
    }

    pub fn to_bool(&self) -> Result<bool> {
        match self {
            Value::Bool(b) => Ok(*b),
            Value::Str(s) => {
                let t = s.trim();
                if t.eq_ignore_ascii_case("true") {
                    Ok(true)
                } else if t.eq_ignore_ascii_case("false") {
                    Ok(false)
                } else {
                    Ok(parse_number(t).ok_or_else(VbError::type_mismatch)? != 0.0)
                }
            }
            Value::Null => Err(VbError::new(94, "Invalid use of Null")),
            other => Ok(other.to_f64()? != 0.0),
        }
    }

    /// The value as text (`CStr`).
    pub fn to_text(&self) -> Result<String> {
        Ok(match self {
            Value::Empty | Value::Missing => String::new(),
            Value::Null => return Err(VbError::new(94, "Invalid use of Null")),
            Value::Bool(b) => if *b { "True" } else { "False" }.to_string(),
            Value::Byte(n) => n.to_string(),
            Value::Int(n) => n.to_string(),
            Value::Long(n) => n.to_string(),
            Value::Single(n) => number_text(f64::from(*n), 7),
            Value::Double(n) => number_text(*n, 15),
            Value::Date(n) => date_text(*n),
            Value::Str(s) => s.to_string(),
            Value::Error(n) => format!("Error {n}"),
            Value::Array(_) => return Err(VbError::type_mismatch()),
            Value::Object(_) => return Err(VbError::new(438, "Object doesn't support this property or method")),
        })
    }

    pub fn to_rc_str(&self) -> Result<Rc<str>> {
        match self {
            Value::Str(s) => Ok(s.clone()),
            other => Ok(Rc::from(other.to_text()?)),
        }
    }
}

/// Rounds half to even (VBA's `CLng`, `CInt` and `Round`).
pub fn round_even(n: f64) -> f64 {
    let r = n.round();
    if (n - n.trunc()).abs() == 0.5 && r % 2.0 != 0.0 { r - n.signum() } else { r }
}

/// Reads text as a number the way VBA does for `CDbl("1,5")`-free input: leading and trailing
/// spaces, a sign, digits, a decimal point, an exponent, `&H` and `&O` literals.
pub fn parse_number(s: &str) -> Option<f64> {
    let t = s.trim();
    if t.is_empty() {
        return None;
    }
    let upper = t.to_ascii_uppercase();
    if let Some(h) = upper.strip_prefix("&H") {
        return i64::from_str_radix(h.trim_end_matches('&'), 16).ok().map(|n| n as f64);
    }
    if let Some(o) = upper.strip_prefix("&O") {
        return i64::from_str_radix(o.trim_end_matches('&'), 8).ok().map(|n| n as f64);
    }
    let t = t.replace(['D', 'd'], "E");
    // `f64::from_str` also accepts "inf" and "nan", which are not VBA numbers.
    if t.chars().any(|c| c.is_ascii_alphabetic() && c != 'e' && c != 'E') {
        return None;
    }
    t.parse::<f64>().ok().filter(|n| n.is_finite())
}

/// `Val`: the number at the start of `s` (0 when there is none), reading no further than it.
pub fn val_prefix(s: &str) -> f64 {
    let t: String = s.chars().filter(|c| !c.is_whitespace()).collect();
    let bytes = t.as_bytes();
    let mut end = 0;
    let mut seen_digit = false;
    let mut seen_dot = false;
    let mut i = 0;
    if matches!(bytes.first(), Some(b'+') | Some(b'-')) {
        i = 1;
    }
    while let Some(&b) = bytes.get(i) {
        if b.is_ascii_digit() {
            seen_digit = true;
            end = i + 1;
        } else if b == b'.' && !seen_dot {
            seen_dot = true;
            if seen_digit {
                end = i + 1;
            }
        } else {
            break;
        }
        i += 1;
    }
    if seen_digit
        && let Some(&e) = bytes.get(end)
        && (e == b'e' || e == b'E' || e == b'd' || e == b'D')
    {
        let mut j = end + 1;
        if matches!(bytes.get(j), Some(b'+') | Some(b'-')) {
            j += 1;
        }
        let start = j;
        while bytes.get(j).is_some_and(u8::is_ascii_digit) {
            j += 1;
        }
        if j > start {
            end = j;
        }
    }
    let part = t.get(..end).unwrap_or("").replace(['d', 'D'], "e");
    part.parse::<f64>().unwrap_or(0.0)
}

/// A number as `CStr` shows it: up to `digits` significant digits, scientific notation for very
/// large and very small values.
pub fn number_text(n: f64, digits: usize) -> String {
    if n == 0.0 {
        return "0".into();
    }
    if !n.is_finite() {
        return "1.#INF".into();
    }
    let sci = format!("{:.*e}", digits.saturating_sub(1), n.abs());
    let (mant, exp) = sci.split_once('e').unwrap_or((sci.as_str(), "0"));
    let exp: i32 = exp.parse().unwrap_or(0);
    let digits_only: String = mant.chars().filter(char::is_ascii_digit).collect();
    let digits_only = digits_only.trim_end_matches('0');
    let digits_only = if digits_only.is_empty() { "0" } else { digits_only };
    let sign = if n < 0.0 { "-" } else { "" };
    if !(-5..digits as i32).contains(&exp) {
        let (first, rest) = digits_only.split_at(1);
        let m = if rest.is_empty() { first.to_string() } else { format!("{first}.{rest}") };
        return format!("{sign}{m}E{}{:02}", if exp < 0 { '-' } else { '+' }, exp.abs());
    }
    if exp >= 0 {
        let int_len = exp as usize + 1;
        let mut s = String::new();
        for i in 0..int_len {
            s.push(digits_only.as_bytes().get(i).map_or('0', |b| *b as char));
        }
        if digits_only.len() > int_len {
            s.push('.');
            s.push_str(digits_only.get(int_len..).unwrap_or(""));
        }
        format!("{sign}{s}")
    } else {
        let zeros = "0".repeat((-exp - 1) as usize);
        format!("{sign}0.{zeros}{digits_only}")
    }
}

/// A date serial as text: `yyyy-mm-dd`, with ` hh:nn:ss` when it has a time of day.
pub fn date_text(serial: f64) -> String {
    use celdas_core::date::datetime_from_serial;
    match datetime_from_serial(celdas_core::DateSystem::D1900, serial) {
        Some(dt) => {
            let has_time = dt.hour != 0 || dt.minute != 0 || dt.second != 0;
            if serial.trunc() == 0.0 {
                format!("{:02}:{:02}:{:02}", dt.hour, dt.minute, dt.second)
            } else if has_time {
                format!("{:04}-{:02}-{:02} {:02}:{:02}:{:02}", dt.year, dt.month, dt.day, dt.hour, dt.minute, dt.second)
            } else {
                format!("{:04}-{:02}-{:02}", dt.year, dt.month, dt.day)
            }
        }
        None => number_text(serial, 15),
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn numbers_print_like_cstr() {
        assert_eq!(number_text(0.1 + 0.2, 15), "0.3");
        assert_eq!(number_text(1.0 / 3.0, 15), "0.333333333333333");
        assert_eq!(number_text(1e-6, 15), "1E-06");
        assert_eq!(number_text(100000.0, 15), "100000");
        assert_eq!(number_text(-2.5, 15), "-2.5");
        assert_eq!(number_text(1.5e20, 15), "1.5E+20");
        assert_eq!(number_text(0.0001, 15), "0.0001");
        assert_eq!(number_text(123456.789, 15), "123456.789");
    }

    #[test]
    fn rounding_is_half_to_even() {
        assert_eq!(round_even(2.5), 2.0);
        assert_eq!(round_even(3.5), 4.0);
        assert_eq!(round_even(-2.5), -2.0);
        assert_eq!(round_even(2.4), 2.0);
        assert_eq!(Value::Double(0.5).to_i32(), Ok(0));
    }

    #[test]
    fn val_reads_the_leading_number() {
        assert_eq!(val_prefix(" 12abc"), 12.0);
        assert_eq!(val_prefix("1.5e3x"), 1500.0);
        assert_eq!(val_prefix("abc"), 0.0);
        assert_eq!(val_prefix("-.5"), -0.5);
    }

    #[test]
    fn text_to_number_rejects_words() {
        assert_eq!(parse_number("inf"), None);
        assert_eq!(parse_number(" 3.5 "), Some(3.5));
        assert_eq!(parse_number("&HFF"), Some(255.0));
        assert_eq!(parse_number("1d2"), Some(100.0));
    }
}
