//! VBA's own functions (`Left`, `Format`, `UBound`, `DateAdd`…) and its named constants.

use celdas_core::date::{self, DateSystem};

use crate::host::CallArg;
use crate::interp::Vm;
use crate::value::{round_even, val_prefix, Result, VArray, Value, VbError};

/// The value of a VBA constant (`vbCrLf`, `vbYes`, `True`…), by lower-case name.
pub fn vb_constant(name: &str) -> Option<Value> {
    let long = |n: i32| Some(Value::Long(n));
    match name {
        "true" => Some(Value::Bool(true)),
        "false" => Some(Value::Bool(false)),
        "vbcrlf" | "vbnewline" => Some(Value::str("\r\n")),
        "vbcr" => Some(Value::str("\r")),
        "vblf" => Some(Value::str("\n")),
        "vbtab" => Some(Value::str("\t")),
        "vbnullstring" | "vbnullchar" => Some(Value::str("")),
        "vbbinarycompare" | "vbok" | "vbokonly" | "vbgeneralnumber" | "vbempty" | "vbsunday" | "vbusesystem" => long(0),
        "vbtextcompare" | "vbcancel" | "vbokcancel" | "vbnull" | "vbmonday" | "vbusesystemdayofweek" | "vbcurrency" => long(1),
        "vbabort" | "vbabortretryignore" | "vbinteger" | "vbtuesday" | "vbfirstjan1" => long(2),
        "vbretry" | "vbyesnocancel" | "vblong" | "vbwednesday" | "vbfirstfourdays" => long(3),
        "vbignore" | "vbyesno" | "vbsingle" | "vbthursday" | "vbfirstfullweek" => long(4),
        "vbyes" | "vbretrycancel" | "vbdouble" | "vbfriday" => long(5),
        "vbno" | "vbdate" | "vbsaturday" => long(7),
        "vbstring" => long(8),
        "vbobject" => long(9),
        "vberror" => long(10),
        "vbboolean" => long(11),
        "vbvariant" => long(12),
        "vbbyte" => long(17),
        "vbarray" => long(8192),
        "vbcritical" => long(16),
        "vbquestion" => long(32),
        "vbexclamation" => long(48),
        "vbinformation" => long(64),
        "vbdefaultbutton2" => long(256),
        "vbdefaultbutton3" => long(512),
        "vbnormal" => long(0),
        "vbreadonly" => long(1),
        "vbhidden" => long(2),
        "vbsystem" => long(4),
        "vbdirectory" => long(16),
        "vbarchive" => long(32),
        _ => None,
    }
}

fn opt(v: &[Value], i: usize) -> Option<&Value> {
    v.get(i).filter(|x| !matches!(x, Value::Missing))
}

fn need(v: &[Value], i: usize) -> Result<&Value> {
    opt(v, i).ok_or_else(|| VbError::new(449, "Argument not optional"))
}

fn num(v: &[Value], i: usize) -> Result<f64> {
    need(v, i)?.to_f64()
}

fn text(v: &[Value], i: usize) -> Result<String> {
    need(v, i)?.to_text()
}

fn int_or(v: &[Value], i: usize, default: i32) -> Result<i32> {
    opt(v, i).map_or(Ok(default), |x| x.to_i32())
}

fn finite(n: f64) -> Result<Value> {
    if n.is_finite() { Ok(Value::Double(n)) } else { Err(VbError::overflow()) }
}

/// A number result that keeps the integer type of an integer argument.
fn same_kind(arg: &Value, n: f64) -> Value {
    match arg {
        Value::Int(_) if n.abs() <= f64::from(i16::MAX) => Value::Int(n as i16),
        Value::Long(_) if n.abs() <= f64::from(i32::MAX) => Value::Long(n as i32),
        Value::Byte(_) if (0.0..=255.0).contains(&n) => Value::Byte(n as u8),
        Value::Single(_) => Value::Single(n as f32),
        _ => Value::Double(n),
    }
}

impl Vm<'_> {
    /// A built-in function by lower-case name, or `None` when it is not one.
    pub(crate) fn builtin(&mut self, name: &str, args: &[CallArg]) -> Option<Result<Value>> {
        let mut v = Vec::with_capacity(args.len());
        for a in args {
            match self.rvalue_arg(a.value.clone()) {
                Ok(x) => v.push(x),
                Err(e) => return Some(Err(e)),
            }
        }
        if let Some(r) = math(self, name, &v) {
            return Some(r);
        }
        if let Some(r) = convert(name, &v) {
            return Some(r);
        }
        if let Some(r) = strings(name, &v) {
            return Some(r);
        }
        if let Some(r) = info(name, &v) {
            return Some(r);
        }
        if let Some(r) = dates(self, name, &v) {
            return Some(r);
        }
        self.interaction(name, &v)
    }

    /// Arrays and `Nothing` pass through untouched; other objects give their default value.
    fn rvalue_arg(&mut self, v: Value) -> Result<Value> {
        match v {
            Value::Object(None) | Value::Array(_) => Ok(v),
            Value::Object(Some(_)) => self.rvalue(v).or_else(|_| Ok(Value::Empty)),
            other => Ok(other),
        }
    }

    fn interaction(&mut self, name: &str, v: &[Value]) -> Option<Result<Value>> {
        Some(match name {
            "msgbox" => self.msg_box(v),
            "inputbox" => {
                let prompt = text(v, 0).unwrap_or_default();
                let title = opt(v, 1).map(Value::to_text).transpose().map(Option::unwrap_or_default);
                let default = opt(v, 2).map(Value::to_text).transpose().map(Option::unwrap_or_default);
                match (title, default) {
                    (Ok(t), Ok(d)) => Ok(Value::string(self.host.input_box(&prompt, &t, &d).unwrap_or_default())),
                    (Err(e), _) | (_, Err(e)) => Err(e),
                }
            }
            "createobject" => text(v, 0).and_then(|p| self.new_object(&p.to_lowercase().replace("scripting.", ""))),
            "getobject" => Err(VbError::new(429, "ActiveX component can't create object")),
            "freefile" => Ok(Value::Long((1..=255).find(|n| !self.files.contains(n)).unwrap_or(1))),
            "dir" => {
                let pattern = opt(v, 0).map(Value::to_text).transpose();
                let attrs = int_or(v, 1, 0);
                match (pattern, attrs) {
                    (Ok(p), Ok(a)) => self.host.dir(&p.unwrap_or_default(), a).map(Value::string),
                    (Err(e), _) | (_, Err(e)) => Err(e),
                }
            }
            "mkdir" => text(v, 0).and_then(|p| self.host.make_dir(&p)).map(|()| Value::Empty),
            "kill" => text(v, 0).and_then(|p| self.host.remove_file(&p)).map(|()| Value::Empty),
            "environ" | "environ$" => text(v, 0).map(|n| Value::string(self.host.environ(&n))),
            "doevents" | "chdir" | "chdrive" | "beep" | "randomize" => {
                if name == "randomize" {
                    self.rng ^= 0x9E37_79B9_7F4A_7C15 ^ self.host.now().to_bits();
                }
                Ok(Value::Empty)
            }
            "rnd" => {
                self.rng ^= self.rng << 13;
                self.rng ^= self.rng >> 7;
                self.rng ^= self.rng << 17;
                self.last_rnd = (self.rng >> 11) as f64 / (1u64 << 53) as f64;
                Ok(Value::Single(self.last_rnd as f32))
            }
            "timer" => {
                let now = self.host.now();
                Ok(Value::Single(((now - now.floor()) * 86_400.0) as f32))
            }
            _ => return None,
        })
    }

    fn msg_box(&mut self, v: &[Value]) -> Result<Value> {
        let prompt = text(v, 0)?;
        let buttons = int_or(v, 1, 0)?;
        let title = match opt(v, 2) {
            Some(t) => t.to_text()?,
            None => String::new(),
        };
        Ok(Value::Long(self.host.msg_box(&prompt, buttons, &title)))
    }
}

fn math(vm: &mut Vm<'_>, name: &str, v: &[Value]) -> Option<Result<Value>> {
    let _ = vm;
    Some(match name {
        "abs" => need(v, 0).and_then(|a| a.to_f64().map(|n| same_kind(a, n.abs()))),
        "sqr" => num(v, 0).and_then(|n| if n < 0.0 { Err(VbError::invalid_call()) } else { Ok(Value::Double(n.sqrt())) }),
        "sin" => num(v, 0).map(|n| Value::Double(n.sin())),
        "cos" => num(v, 0).map(|n| Value::Double(n.cos())),
        "tan" => num(v, 0).map(|n| Value::Double(n.tan())),
        "atn" => num(v, 0).map(|n| Value::Double(n.atan())),
        "exp" => num(v, 0).and_then(|n| finite(n.exp())),
        "log" => num(v, 0).and_then(|n| if n <= 0.0 { Err(VbError::invalid_call()) } else { Ok(Value::Double(n.ln())) }),
        "int" => need(v, 0).and_then(|a| a.to_f64().map(|n| same_kind(a, n.floor()))),
        "fix" => need(v, 0).and_then(|a| a.to_f64().map(|n| same_kind(a, n.trunc()))),
        "sgn" => num(v, 0).map(|n| Value::Int(if n > 0.0 { 1 } else if n < 0.0 { -1 } else { 0 })),
        "round" => round(v),
        _ => return None,
    })
}

fn round(v: &[Value]) -> Result<Value> {
    let n = num(v, 0)?;
    let digits = int_or(v, 1, 0)?;
    if !(0..=15).contains(&digits) {
        return Err(VbError::invalid_call());
    }
    let scale = 10f64.powi(digits);
    Ok(Value::Double(round_even(n * scale) / scale))
}

fn convert(name: &str, v: &[Value]) -> Option<Result<Value>> {
    use crate::ast::VType;
    let to = |ty: VType| need(v, 0).and_then(|a| crate::ops::coerce(&ty, a.clone()));
    Some(match name {
        "cint" => to(VType::Integer),
        "clng" => to(VType::Long),
        "cdbl" | "ccur" | "cdec" => to(VType::Double),
        "csng" => to(VType::Single),
        "cstr" => to(VType::String),
        "cbool" => to(VType::Boolean),
        "cbyte" => to(VType::Byte),
        "cdate" | "cvdate" => to(VType::Date),
        "cvar" => need(v, 0).cloned(),
        "val" => text(v, 0).map(|s| Value::Double(val_prefix(&s))),
        "str" | "str$" => num(v, 0).and_then(|n| {
            let t = Value::Double(n).to_text()?;
            Ok(Value::string(if n >= 0.0 { format!(" {t}") } else { t }))
        }),
        "hex" | "hex$" => need(v, 0).and_then(|a| a.to_i64()).map(|n| Value::string(format!("{:X}", n as i32))),
        "oct" | "oct$" => need(v, 0).and_then(|a| a.to_i64()).map(|n| Value::string(format!("{:o}", n as i32))),
        "chr" | "chr$" | "chrw" | "chrw$" => num(v, 0).and_then(|n| {
            let code = u32::try_from(n as i64).map_err(|_| VbError::invalid_call())?;
            char::from_u32(code).map(|c| Value::string(c.to_string())).ok_or_else(VbError::invalid_call)
        }),
        "asc" | "ascw" => text(v, 0).and_then(|s| s.chars().next().map(|c| Value::Long(c as i32)).ok_or_else(VbError::invalid_call)),
        _ => return None,
    })
}

fn strings(name: &str, v: &[Value]) -> Option<Result<Value>> {
    Some(match name {
        "len" | "lenb" => need(v, 0).and_then(|a| match a {
            Value::Null => Ok(Value::Null),
            other => other.to_text().map(|s| Value::Long(s.chars().count() as i32)),
        }),
        "left" | "left$" => left_right(v, true),
        "right" | "right$" => left_right(v, false),
        "mid" | "mid$" => mid(v),
        "ucase" | "ucase$" => text(v, 0).map(|s| Value::string(s.to_uppercase())),
        "lcase" | "lcase$" => text(v, 0).map(|s| Value::string(s.to_lowercase())),
        "trim" | "trim$" => text(v, 0).map(|s| Value::string(s.trim_matches(' ').to_string())),
        "ltrim" | "ltrim$" => text(v, 0).map(|s| Value::string(s.trim_start_matches(' ').to_string())),
        "rtrim" | "rtrim$" => text(v, 0).map(|s| Value::string(s.trim_end_matches(' ').to_string())),
        "space" | "space$" => count(v, 0).map(|n| Value::string(" ".repeat(n))),
        "string" | "string$" => string_fn(v),
        "strreverse" => text(v, 0).map(|s| Value::string(s.chars().rev().collect())),
        "instr" => instr(v),
        "instrrev" => instr_rev(v),
        "replace" | "replace$" => replace(v),
        "strcomp" => str_comp(v),
        "split" => split(v),
        "join" => join(v),
        "format" | "format$" => need(v, 0).and_then(|a| {
            let code = opt(v, 1).map(Value::to_text).transpose()?.unwrap_or_default();
            crate::format::format(a, &code).map(Value::string)
        }),
        "formatnumber" => format_fixed(v, false),
        "formatpercent" => format_fixed(v, true),
        _ => return None,
    })
}

/// A length argument: non-negative, and small enough not to exhaust memory.
fn count(v: &[Value], i: usize) -> Result<usize> {
    let n = need(v, i)?.to_i32()?;
    if !(0..=10_000_000).contains(&n) {
        return Err(VbError::invalid_call());
    }
    Ok(n as usize)
}

fn left_right(v: &[Value], left: bool) -> Result<Value> {
    let s = text(v, 0)?;
    let n = need(v, 1)?.to_i32()?;
    if n < 0 {
        return Err(VbError::invalid_call());
    }
    let n = n as usize;
    let len = s.chars().count();
    let out: String = if left { s.chars().take(n).collect() } else { s.chars().skip(len.saturating_sub(n)).collect() };
    Ok(Value::string(out))
}

fn mid(v: &[Value]) -> Result<Value> {
    let s = text(v, 0)?;
    let start = need(v, 1)?.to_i32()?;
    if start < 1 {
        return Err(VbError::invalid_call());
    }
    let skip = (start - 1) as usize;
    let out: String = match opt(v, 2) {
        Some(l) => {
            let l = l.to_i32()?;
            if l < 0 {
                return Err(VbError::invalid_call());
            }
            s.chars().skip(skip).take(l as usize).collect()
        }
        None => s.chars().skip(skip).collect(),
    };
    Ok(Value::string(out))
}

fn string_fn(v: &[Value]) -> Result<Value> {
    let n = count(v, 0)?;
    let c = match need(v, 1)? {
        Value::Str(s) => s.chars().next(),
        other => char::from_u32(other.to_i32()?.unsigned_abs()),
    };
    let c = c.ok_or_else(VbError::invalid_call)?;
    Ok(Value::string(std::iter::repeat_n(c, n).collect()))
}

/// Where `needle` first occurs in `hay` at or after char index `from`, comparing with `fold`.
fn find_from(hay: &[char], needle: &[char], from: usize, fold: bool) -> Option<usize> {
    if needle.is_empty() {
        return Some(from);
    }
    let same = |a: char, b: char| if fold { a.to_lowercase().eq(b.to_lowercase()) } else { a == b };
    (from..hay.len().checked_sub(needle.len())? + 1).find(|&i| needle.iter().enumerate().all(|(k, &c)| hay.get(i + k).is_some_and(|&h| same(h, c))))
}

fn instr(v: &[Value]) -> Result<Value> {
    // InStr([start,] hay, needle [, compare]): a numeric first argument is the start.
    let has_start = matches!(v.first(), Some(a) if !matches!(a, Value::Str(_)) && v.len() >= 3) || (v.len() == 4);
    let (start, rest) = if has_start { (need(v, 0)?.to_i32()?, v.get(1..).unwrap_or(&[])) } else { (1, v) };
    if start < 1 {
        return Err(VbError::invalid_call());
    }
    let hay: Vec<char> = text(rest, 0)?.chars().collect();
    let needle: Vec<char> = text(rest, 1)?.chars().collect();
    let fold = int_or(rest, 2, 0)? == 1;
    let from = (start - 1) as usize;
    if from > hay.len() {
        return Ok(Value::Long(0));
    }
    Ok(Value::Long(find_from(&hay, &needle, from, fold).map_or(0, |i| i as i32 + 1)))
}

fn instr_rev(v: &[Value]) -> Result<Value> {
    let hay: Vec<char> = text(v, 0)?.chars().collect();
    let needle: Vec<char> = text(v, 1)?.chars().collect();
    let start = int_or(v, 2, -1)?;
    let fold = int_or(v, 3, 0)? == 1;
    let last = if start < 0 { hay.len() } else { (start as usize).min(hay.len()) };
    if needle.is_empty() {
        return Ok(Value::Long(last as i32));
    }
    let mut best = 0;
    let mut from = 0;
    while let Some(i) = find_from(&hay, &needle, from, fold) {
        if i + needle.len() > last {
            break;
        }
        best = i as i32 + 1;
        from = i + 1;
    }
    Ok(Value::Long(best))
}

fn replace(v: &[Value]) -> Result<Value> {
    let s = text(v, 0)?;
    let find: Vec<char> = text(v, 1)?.chars().collect();
    let with = text(v, 2)?;
    let start = int_or(v, 3, 1)?;
    let limit = int_or(v, 4, -1)?;
    let fold = int_or(v, 5, 0)? == 1;
    if start < 1 {
        return Err(VbError::invalid_call());
    }
    let hay: Vec<char> = s.chars().collect();
    let skip = ((start - 1) as usize).min(hay.len());
    if find.is_empty() {
        return Ok(Value::string(hay.iter().skip(skip).collect()));
    }
    let mut out = String::new();
    let mut i = skip;
    let mut done = 0;
    while let Some(at) = find_from(&hay, &find, i, fold) {
        if limit >= 0 && done >= limit {
            break;
        }
        out.extend(hay.get(i..at).unwrap_or(&[]));
        out.push_str(&with);
        i = at + find.len();
        done += 1;
    }
    out.extend(hay.get(i..).unwrap_or(&[]));
    Ok(Value::string(out))
}

fn str_comp(v: &[Value]) -> Result<Value> {
    let (a, b) = (text(v, 0)?, text(v, 1)?);
    let ord = if int_or(v, 2, 0)? == 1 { a.to_lowercase().cmp(&b.to_lowercase()) } else { a.cmp(&b) };
    Ok(Value::Int(ord as i16))
}

fn split(v: &[Value]) -> Result<Value> {
    let s = text(v, 0)?;
    let delim = match opt(v, 1) {
        Some(d) => d.to_text()?,
        None => " ".to_string(),
    };
    let limit = int_or(v, 2, -1)?;
    let fold = int_or(v, 3, 0)? == 1;
    let parts = split_text(&s, &delim, limit, fold);
    Ok(Value::array(VArray::from_vec(parts.into_iter().map(Value::string).collect(), 0)))
}

fn split_text(s: &str, delim: &str, limit: i32, fold: bool) -> Vec<String> {
    if s.is_empty() {
        return Vec::new();
    }
    if delim.is_empty() || limit == 1 {
        return vec![s.to_string()];
    }
    let hay: Vec<char> = s.chars().collect();
    let needle: Vec<char> = delim.chars().collect();
    let mut out = Vec::new();
    let mut from = 0;
    while limit < 0 || (out.len() as i32) < limit - 1 {
        let Some(at) = find_from(&hay, &needle, from, fold) else { break };
        out.push(hay.get(from..at).unwrap_or(&[]).iter().collect());
        from = at + needle.len();
    }
    out.push(hay.get(from..).unwrap_or(&[]).iter().collect());
    out
}

fn join(v: &[Value]) -> Result<Value> {
    let Value::Array(a) = need(v, 0)? else { return Err(VbError::type_mismatch()) };
    let delim = match opt(v, 1) {
        Some(d) => d.to_text()?,
        None => " ".to_string(),
    };
    let parts = a.borrow().data.iter().map(Value::to_text).collect::<Result<Vec<_>>>()?;
    Ok(Value::string(parts.join(&delim)))
}

fn format_fixed(v: &[Value], percent: bool) -> Result<Value> {
    let n = num(v, 0)?;
    let digits = int_or(v, 1, 2)?.clamp(0, 15) as usize;
    let code = format!("{}{}{}", if digits > 0 { "#,##0." } else { "#,##0" }, "0".repeat(digits), if percent { "%" } else { "" });
    crate::format::format(&Value::Double(n), &code).map(Value::string)
}

fn info(name: &str, v: &[Value]) -> Option<Result<Value>> {
    let first = || need(v, 0);
    let flag = |f: &dyn Fn(&Value) -> bool| first().map(|a| Value::Bool(f(a)));
    Some(match name {
        "isnumeric" => flag(&|a| match a {
            Value::Str(s) => crate::value::parse_number(s).is_some(),
            Value::Bool(_) | Value::Empty | Value::Null | Value::Date(_) | Value::Error(_) | Value::Array(_) | Value::Object(_) | Value::Missing => false,
            _ => true,
        }),
        "isdate" => flag(&|a| match a {
            Value::Date(_) => true,
            Value::Str(s) => crate::lexer::parse_date_literal(s).is_some(),
            _ => false,
        }),
        "isempty" => flag(&|a| matches!(a, Value::Empty | Value::Missing)),
        "isnull" => flag(&|a| matches!(a, Value::Null)),
        "iserror" => flag(&|a| matches!(a, Value::Error(_))),
        "isarray" => flag(&|a| matches!(a, Value::Array(_))),
        "isobject" => flag(&|a| matches!(a, Value::Object(_))),
        "ismissing" => Ok(Value::Bool(matches!(v.first(), Some(Value::Missing) | None))),
        "typename" => first().map(|a| Value::string(a.type_name())),
        "vartype" => first().map(|a| Value::Int(a.var_type() as i16)),
        "array" => Ok(Value::array(VArray::from_vec(v.to_vec(), 0))),
        "lbound" => bound(v, true),
        "ubound" => bound(v, false),
        "iif" => iif(v),
        "choose" => choose(v),
        "rgb" => rgb(v),
        "cverr" => first().and_then(|a| a.to_i32()).map(Value::Error),
        _ => return None,
    })
}

fn bound(v: &[Value], lower: bool) -> Result<Value> {
    let Value::Array(a) = need(v, 0)? else { return Err(VbError::type_mismatch()) };
    let dim = int_or(v, 1, 1)?;
    let a = a.borrow();
    let idx = usize::try_from(dim - 1).map_err(|_| VbError::subscript())?;
    let b = a.bounds.get(idx).ok_or_else(VbError::subscript)?;
    Ok(Value::Long(if lower { b.0 } else { b.1 }))
}

fn iif(v: &[Value]) -> Result<Value> {
    let pick = if need(v, 0)?.to_bool()? { 1 } else { 2 };
    need(v, pick).cloned()
}

fn choose(v: &[Value]) -> Result<Value> {
    let n = need(v, 0)?.to_i32()?;
    match usize::try_from(n) {
        Ok(i) if i >= 1 => Ok(v.get(i).cloned().unwrap_or(Value::Null)),
        _ => Ok(Value::Null),
    }
}

fn rgb(v: &[Value]) -> Result<Value> {
    let c = |i| need(v, i).and_then(Value::to_i32).map(|n| n.clamp(0, 255));
    Ok(Value::Long(c(0)? | (c(1)? << 8) | (c(2)? << 16)))
}

fn dates(vm: &mut Vm<'_>, name: &str, v: &[Value]) -> Option<Result<Value>> {
    Some(match name {
        "now" => Ok(Value::Date(vm.host.now())),
        "date" => Ok(Value::Date(vm.host.now().floor())),
        "time" => {
            let now = vm.host.now();
            Ok(Value::Date(now - now.floor()))
        }
        "year" | "month" | "day" | "hour" | "minute" | "second" | "weekday" => date_part(name, v),
        "dateserial" => date_serial(v),
        "timeserial" => {
            let (h, m, s) = (num(v, 0), num(v, 1), num(v, 2));
            match (h, m, s) {
                (Ok(h), Ok(m), Ok(s)) => Ok(Value::Date(date::time_fraction(h, m, s).rem_euclid(1.0))),
                (Err(e), _, _) | (_, Err(e), _) | (_, _, Err(e)) => Err(e),
            }
        }
        "datevalue" => text(v, 0).and_then(|s| crate::lexer::parse_date_literal(&s).map(|n| Value::Date(n.floor())).ok_or_else(VbError::type_mismatch)),
        "timevalue" => text(v, 0).and_then(|s| crate::lexer::parse_date_literal(&s).map(|n| Value::Date(n - n.floor())).ok_or_else(VbError::type_mismatch)),
        "dateadd" => date_add(v),
        "datediff" => date_diff(v),
        "monthname" => num(v, 0).and_then(|n| month_name(n as i32, opt(v, 1).is_some_and(|a| a.to_bool().unwrap_or(false)))),
        "weekdayname" => num(v, 0).and_then(|n| weekday_name(n as i32, opt(v, 1).is_some_and(|a| a.to_bool().unwrap_or(false)))),
        _ => return None,
    })
}

fn parts(serial: f64) -> Result<date::DateTime> {
    date::datetime_from_serial(DateSystem::D1900, serial).ok_or_else(VbError::invalid_call)
}

fn date_part(name: &str, v: &[Value]) -> Result<Value> {
    let serial = match need(v, 0)? {
        Value::Str(s) => crate::lexer::parse_date_literal(s).ok_or_else(VbError::type_mismatch)?,
        other => other.to_f64()?,
    };
    let p = parts(serial)?;
    Ok(Value::Long(match name {
        "year" => p.year,
        "month" => p.month as i32,
        "day" => p.day as i32,
        "hour" => p.hour as i32,
        "minute" => p.minute as i32,
        "second" => p.second as i32,
        _ => {
            // Weekday(date, firstDayOfWeek): 1 = Sunday by default.
            let first = int_or(v, 1, 1)?.clamp(1, 7);
            let first = if first == 0 { 1 } else { first };
            ((p.weekday as i32 + 7 - (first - 1)) % 7) + 1
        }
    }))
}

fn date_serial(v: &[Value]) -> Result<Value> {
    let (y, m, d) = (need(v, 0)?.to_i64()?, need(v, 1)?.to_i64()?, need(v, 2)?.to_i64()?);
    date::serial_from_ymd(DateSystem::D1900, y, m, d).map(Value::Date).ok_or_else(VbError::invalid_call)
}

fn date_add(v: &[Value]) -> Result<Value> {
    let unit = text(v, 0)?.to_lowercase();
    let n = need(v, 1)?.to_f64()?.trunc();
    let serial = need(v, 2)?.to_f64()?;
    let out = match unit.as_str() {
        "d" | "y" | "w" => serial + n,
        "ww" => serial + 7.0 * n,
        "h" => serial + n / 24.0,
        "n" => serial + n / 1440.0,
        "s" => serial + n / 86_400.0,
        "m" | "q" | "yyyy" => {
            let months = match unit.as_str() {
                "q" => n * 3.0,
                "yyyy" => n * 12.0,
                _ => n,
            };
            let p = parts(serial)?;
            let target = i64::from(p.month) - 1 + months as i64;
            let (year, month) = (i64::from(p.year) + target.div_euclid(12), target.rem_euclid(12) + 1);
            let last = i64::from(date::days_in_month(year as i32, month as u32));
            let day = i64::from(p.day).min(last);
            let base = date::serial_from_ymd(DateSystem::D1900, year, month, day).ok_or_else(VbError::overflow)?;
            base + (serial - serial.floor())
        }
        _ => return Err(VbError::invalid_call()),
    };
    Ok(Value::Date(out))
}

fn date_diff(v: &[Value]) -> Result<Value> {
    let unit = text(v, 0)?.to_lowercase();
    let (a, b) = (num(v, 1)?, num(v, 2)?);
    let n = match unit.as_str() {
        "d" | "y" => b.floor() - a.floor(),
        "w" | "ww" => ((b.floor() - a.floor()) / 7.0).trunc(),
        "h" => ((b - a) * 24.0).trunc(),
        "n" => ((b - a) * 1440.0).trunc(),
        "s" => ((b - a) * 86_400.0).round(),
        "m" | "q" | "yyyy" => {
            let (pa, pb) = (parts(a)?, parts(b)?);
            let months = f64::from((pb.year - pa.year) * 12 + pb.month as i32 - pa.month as i32);
            match unit.as_str() {
                "q" => (months / 3.0).trunc(),
                "yyyy" => f64::from(pb.year - pa.year),
                _ => months,
            }
        }
        _ => return Err(VbError::invalid_call()),
    };
    Ok(Value::Long(n as i32))
}

const MONTHS: [&str; 12] = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
const DAYS: [&str; 7] = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"];

fn month_name(n: i32, short: bool) -> Result<Value> {
    let name = usize::try_from(n - 1).ok().and_then(|i| MONTHS.get(i)).ok_or_else(VbError::invalid_call)?;
    Ok(Value::str(if short { name.get(..3).unwrap_or(name) } else { name }))
}

fn weekday_name(n: i32, short: bool) -> Result<Value> {
    let name = usize::try_from(n - 1).ok().and_then(|i| DAYS.get(i)).ok_or_else(VbError::invalid_call)?;
    Ok(Value::str(if short { name.get(..3).unwrap_or(name) } else { name }))
}
