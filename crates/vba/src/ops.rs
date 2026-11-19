//! Operators and conversions on [`Value`], with VBA's rules for mixed types.

use std::cmp::Ordering;

use crate::ast::{Bin, Un, VType};
use crate::value::{Result, Value, VbError, parse_number, round_even};

/// Converts `v` to the declared type of a variable (what an assignment does).
pub fn coerce(ty: &VType, v: Value) -> Result<Value> {
    if matches!(ty, VType::Variant) {
        return Ok(v);
    }
    if let Value::Object(_) = &v {
        return match ty {
            VType::Object(_) => Ok(v),
            _ => Err(VbError::type_mismatch()),
        };
    }
    if let Value::Array(_) = &v {
        return Err(VbError::type_mismatch());
    }
    Ok(match ty {
        VType::Variant => v,
        VType::Boolean => Value::Bool(v.to_bool()?),
        VType::Byte => {
            let n = v.to_i32()?;
            Value::Byte(u8::try_from(n).map_err(|_| VbError::overflow())?)
        }
        VType::Integer => {
            let n = v.to_i32()?;
            Value::Int(i16::try_from(n).map_err(|_| VbError::overflow())?)
        }
        VType::Long => Value::Long(v.to_i32()?),
        VType::Single => Value::Single(v.to_f64()? as f32),
        VType::Double | VType::Currency => Value::Double(v.to_f64()?),
        VType::Date => match &v {
            Value::Str(s) => Value::Date(crate::lexer::parse_date_literal(s).ok_or_else(VbError::type_mismatch)?),
            _ => Value::Date(v.to_f64()?),
        },
        VType::String => Value::string(v.to_text()?),
        VType::Object(_) => return Err(VbError::type_mismatch()),
    })
}

/// The value a variable of type `ty` holds before it is assigned.
pub fn default_value(ty: &VType) -> Value {
    match ty {
        VType::Variant => Value::Empty,
        VType::Boolean => Value::Bool(false),
        VType::Byte => Value::Byte(0),
        VType::Integer => Value::Int(0),
        VType::Long => Value::Long(0),
        VType::Single => Value::Single(0.0),
        VType::Double | VType::Currency => Value::Double(0.0),
        VType::Date => Value::Date(0.0),
        VType::String => Value::str(""),
        VType::Object(_) => Value::nothing(),
    }
}

fn is_integral(v: &Value) -> bool {
    matches!(v, Value::Byte(_) | Value::Int(_) | Value::Long(_) | Value::Bool(_) | Value::Empty)
}

fn is_null(v: &Value) -> bool {
    matches!(v, Value::Null)
}

/// A number in the narrowest type that holds it, the way VBA sizes the result of arithmetic on
/// integral operands.
fn integral_result(n: f64) -> Value {
    if n.fract() == 0.0 && n.abs() <= 2_147_483_647.0 {
        let n = n as i32;
        match i16::try_from(n) {
            Ok(i) => Value::Int(i),
            Err(_) => Value::Long(n),
        }
    } else {
        Value::Double(n)
    }
}

/// The numeric operand of an arithmetic operator: text must read as a number.
fn operand(v: &Value) -> Result<f64> {
    match v {
        Value::Str(s) => parse_number(s).ok_or_else(VbError::type_mismatch),
        other => other.to_f64(),
    }
}

pub fn arith(op: Bin, a: &Value, b: &Value) -> Result<Value> {
    if is_null(a) || is_null(b) {
        return Ok(Value::Null);
    }
    match op {
        Bin::Add => {
            if let (Value::Str(x), Value::Str(y)) = (a, b) {
                return Ok(Value::string(format!("{x}{y}")));
            }
            let (x, y) = (operand(a)?, operand(b)?);
            Ok(number_result(a, b, x + y))
        }
        Bin::Sub => {
            let (x, y) = (operand(a)?, operand(b)?);
            if matches!((a, b), (Value::Date(_), Value::Date(_))) {
                return Ok(Value::Double(x - y));
            }
            Ok(number_result(a, b, x - y))
        }
        Bin::Mul => {
            let (x, y) = (operand(a)?, operand(b)?);
            Ok(number_result(a, b, x * y))
        }
        Bin::Div => {
            let (x, y) = (operand(a)?, operand(b)?);
            if y == 0.0 {
                return Err(VbError::div_zero());
            }
            Ok(Value::Double(x / y))
        }
        Bin::IntDiv => {
            let (x, y) = (round_even(operand(a)?), round_even(operand(b)?));
            if y == 0.0 {
                return Err(VbError::div_zero());
            }
            let q = (x / y).trunc();
            if q.abs() > 2_147_483_647.0 {
                return Err(VbError::overflow());
            }
            Ok(Value::Long(q as i32))
        }
        Bin::Mod => {
            let (x, y) = (round_even(operand(a)?), round_even(operand(b)?));
            if y == 0.0 {
                return Err(VbError::div_zero());
            }
            Ok(Value::Long((x % y) as i32))
        }
        Bin::Pow => {
            let (x, y) = (operand(a)?, operand(b)?);
            let r = x.powf(y);
            if r.is_nan() || r.is_infinite() {
                return Err(VbError::overflow());
            }
            Ok(Value::Double(r))
        }
        _ => Err(VbError::type_mismatch()),
    }
}

fn number_result(a: &Value, b: &Value, n: f64) -> Value {
    if matches!(a, Value::Date(_)) || matches!(b, Value::Date(_)) {
        return Value::Date(n);
    }
    if is_integral(a) && is_integral(b) {
        return integral_result(n);
    }
    if matches!(a, Value::Single(_)) && matches!(b, Value::Single(_)) {
        return Value::Single(n as f32);
    }
    Value::Double(n)
}

pub fn concat(a: &Value, b: &Value) -> Result<Value> {
    let text = |v: &Value| if is_null(v) { Ok(String::new()) } else { v.to_text() };
    Ok(Value::string(format!("{}{}", text(a)?, text(b)?)))
}

/// Compares two values the way `=`, `<` and friends do. `None` means "not comparable" (Null).
pub fn compare(a: &Value, b: &Value) -> Result<Option<Ordering>> {
    if is_null(a) || is_null(b) {
        return Ok(None);
    }
    match (a, b) {
        (Value::Str(x), Value::Str(y)) => Ok(Some(x.as_ref().cmp(y.as_ref()))),
        (Value::Empty, Value::Str(s)) | (Value::Str(s), Value::Empty) => {
            let o = s.is_empty();
            let ord = if o {
                Ordering::Equal
            } else if matches!(a, Value::Empty) {
                Ordering::Less
            } else {
                Ordering::Greater
            };
            Ok(Some(ord))
        }
        (Value::Str(s), n) if n.as_number().is_some() || matches!(n, Value::Empty) => compare_text_number(s, n, false),
        (n, Value::Str(s)) if n.as_number().is_some() || matches!(n, Value::Empty) => compare_text_number(s, n, true),
        (Value::Error(_), _) | (_, Value::Error(_)) => Err(VbError::type_mismatch()),
        (Value::Object(_), _) | (_, Value::Object(_)) | (Value::Array(_), _) | (_, Value::Array(_)) => Err(VbError::type_mismatch()),
        _ => {
            let x = a.to_f64()?;
            let y = b.to_f64()?;
            Ok(x.partial_cmp(&y))
        }
    }
}

fn compare_text_number(s: &str, n: &Value, flipped: bool) -> Result<Option<Ordering>> {
    let x = parse_number(s).ok_or_else(VbError::type_mismatch)?;
    let y = n.to_f64()?;
    let ord = x.partial_cmp(&y);
    Ok(if flipped { ord.map(Ordering::reverse) } else { ord })
}

pub fn comparison(op: Bin, a: &Value, b: &Value) -> Result<Value> {
    let Some(ord) = compare(a, b)? else {
        return Ok(if is_null(a) || is_null(b) { Value::Null } else { Value::Bool(op == Bin::Ne) });
    };
    let r = match op {
        Bin::Eq => ord == Ordering::Equal,
        Bin::Ne => ord != Ordering::Equal,
        Bin::Lt => ord == Ordering::Less,
        Bin::Gt => ord == Ordering::Greater,
        Bin::Le => ord != Ordering::Greater,
        Bin::Ge => ord != Ordering::Less,
        _ => return Err(VbError::type_mismatch()),
    };
    Ok(Value::Bool(r))
}

pub fn logic(op: Bin, a: &Value, b: &Value) -> Result<Value> {
    if let (Value::Bool(x), Value::Bool(y)) = (a, b) {
        return Ok(Value::Bool(match op {
            Bin::And => *x && *y,
            Bin::Or => *x || *y,
            Bin::Xor => x != y,
            Bin::Eqv => x == y,
            _ => !*x || *y,
        }));
    }
    if is_null(a) || is_null(b) {
        // Null And False is False, Null Or True is True; anything else stays Null.
        return Ok(match (op, a, b) {
            (Bin::And, Value::Bool(false), _) | (Bin::And, _, Value::Bool(false)) => Value::Bool(false),
            (Bin::Or, Value::Bool(true), _) | (Bin::Or, _, Value::Bool(true)) => Value::Bool(true),
            _ => Value::Null,
        });
    }
    let (x, y) = (a.to_i32()?, b.to_i32()?);
    Ok(Value::Long(match op {
        Bin::And => x & y,
        Bin::Or => x | y,
        Bin::Xor => x ^ y,
        Bin::Eqv => !(x ^ y),
        _ => !x | y,
    }))
}

pub fn unary(op: Un, v: &Value) -> Result<Value> {
    match op {
        Un::Plus => Ok(v.clone()),
        Un::Neg => match v {
            Value::Null => Ok(Value::Null),
            Value::Int(n) => Ok(match n.checked_neg() {
                Some(m) => Value::Int(m),
                None => Value::Long(-i32::from(*n)),
            }),
            Value::Long(n) => Ok(match n.checked_neg() {
                Some(m) => Value::Long(m),
                None => Value::Double(-f64::from(*n)),
            }),
            Value::Single(n) => Ok(Value::Single(-n)),
            Value::Bool(b) => Ok(Value::Int(if *b { 1 } else { 0 })),
            Value::Byte(n) => Ok(Value::Int(-i16::from(*n))),
            Value::Date(n) => Ok(Value::Date(-n)),
            other => Ok(Value::Double(-operand(other)?)),
        },
        Un::Not => match v {
            Value::Null => Ok(Value::Null),
            Value::Bool(b) => Ok(Value::Bool(!b)),
            other => Ok(Value::Long(!other.to_i32()?)),
        },
    }
}

/// The `Like` operator: `?` one character, `*` any run, `#` one digit, `[abc]`, `[!a-c]`.
pub fn like(text: &str, pattern: &str) -> bool {
    let t: Vec<char> = text.chars().collect();
    let p: Vec<char> = pattern.chars().collect();
    like_at(&t, 0, &p, 0, 0)
}

fn like_at(t: &[char], ti: usize, p: &[char], pi: usize, depth: u32) -> bool {
    if depth > 200 {
        return false;
    }
    let Some(&pc) = p.get(pi) else {
        return ti == t.len();
    };
    match pc {
        '*' => (ti..=t.len()).any(|k| like_at(t, k, p, pi + 1, depth + 1)),
        '?' => ti < t.len() && like_at(t, ti + 1, p, pi + 1, depth + 1),
        '#' => t.get(ti).is_some_and(char::is_ascii_digit) && like_at(t, ti + 1, p, pi + 1, depth + 1),
        '[' => {
            let Some(close) = p.iter().skip(pi + 1).position(|c| *c == ']').map(|k| k + pi + 1) else {
                return t.get(ti) == Some(&'[') && like_at(t, ti + 1, p, pi + 1, depth + 1);
            };
            let Some(&c) = t.get(ti) else { return false };
            let mut set = p.get(pi + 1..close).unwrap_or(&[]);
            let negate = set.first() == Some(&'!');
            if negate {
                set = set.get(1..).unwrap_or(&[]);
            }
            let mut hit = false;
            let mut i = 0;
            while let Some(&lo) = set.get(i) {
                if set.get(i + 1) == Some(&'-')
                    && let Some(&hi) = set.get(i + 2)
                {
                    hit |= (lo..=hi).contains(&c);
                    i += 3;
                } else {
                    hit |= lo == c;
                    i += 1;
                }
            }
            hit != negate && like_at(t, ti + 1, p, close + 1, depth + 1)
        }
        c => t.get(ti) == Some(&c) && like_at(t, ti + 1, p, pi + 1, depth + 1),
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn integral_arithmetic_stays_integral() {
        assert_eq!(arith(Bin::Add, &Value::Int(2), &Value::Int(3)).unwrap(), Value::Int(5));
        assert!(matches!(arith(Bin::Add, &Value::Int(30000), &Value::Int(30000)).unwrap(), Value::Long(60000)));
        assert!(matches!(arith(Bin::Div, &Value::Int(1), &Value::Int(2)).unwrap(), Value::Double(_)));
        assert_eq!(arith(Bin::Div, &Value::Int(1), &Value::Int(0)), Err(VbError::div_zero()));
    }

    #[test]
    fn text_and_numbers() {
        assert_eq!(arith(Bin::Add, &Value::str("1"), &Value::Int(2)).unwrap().to_f64().unwrap(), 3.0);
        assert_eq!(arith(Bin::Add, &Value::str("a"), &Value::str("b")).unwrap(), Value::str("ab"));
        assert!(arith(Bin::Add, &Value::str("a"), &Value::Int(1)).is_err());
        assert_eq!(comparison(Bin::Eq, &Value::Empty, &Value::str("")).unwrap(), Value::Bool(true));
        assert_eq!(comparison(Bin::Lt, &Value::str("a"), &Value::str("b")).unwrap(), Value::Bool(true));
    }

    #[test]
    fn mod_and_int_division_round_first() {
        assert_eq!(arith(Bin::IntDiv, &Value::Double(7.6), &Value::Int(2)).unwrap(), Value::Long(4));
        assert_eq!(arith(Bin::Mod, &Value::Int(10), &Value::Int(3)).unwrap(), Value::Long(1));
    }

    #[test]
    fn like_patterns() {
        assert!(like("abc123", "abc###"));
        assert!(like("hello", "h*o"));
        assert!(like("cat", "[a-c]at"));
        assert!(!like("dat", "[!d]at"));
        assert!(like("", "*"));
        assert!(!like("ab", "a"));
    }

    #[test]
    fn coercion_follows_the_declared_type() {
        assert_eq!(coerce(&VType::Long, Value::Double(2.5)).unwrap(), Value::Long(2));
        assert_eq!(coerce(&VType::Integer, Value::Long(40000)), Err(VbError::overflow()));
        assert_eq!(coerce(&VType::String, Value::Double(1.5)).unwrap(), Value::str("1.5"));
    }
}
