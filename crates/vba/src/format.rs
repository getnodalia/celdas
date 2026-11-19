//! VBA's `Format` function, on top of the spreadsheet number formats.

use celdas_core::{DateSystem, Value as CoreValue};

use crate::value::{parse_number, Result, Value, VbError};

/// `Format(value, code)`. Decimal and thousands separators are the English ones.
pub fn format(v: &Value, code: &str) -> Result<String> {
    if matches!(v, Value::Null) {
        return Ok(String::new());
    }
    let named = named_format(&code.to_ascii_lowercase());
    let code = match &named {
        Some(Named::Code(c)) => *c,
        Some(Named::General) => return v.to_text(),
        Some(Named::Words(yes, no)) => {
            let on = v.to_bool()?;
            return Ok((if on { *yes } else { *no }).to_string());
        }
        None => code,
    };
    if code.is_empty() {
        return v.to_text();
    }
    let code = minutes_as_spreadsheet(code);
    let core = match v {
        Value::Str(s) => match parse_number(s) {
            Some(n) => CoreValue::number(n),
            None => CoreValue::Text(s.as_ref().into()),
        },
        Value::Bool(b) => CoreValue::number(if *b { -1.0 } else { 0.0 }),
        Value::Empty | Value::Missing => CoreValue::number(0.0),
        other => CoreValue::number(other.to_f64()?),
    };
    celdas_numfmt::text_function(&core, &code, DateSystem::D1900).map_err(|_| VbError::invalid_call())
}

enum Named {
    Code(&'static str),
    General,
    Words(&'static str, &'static str),
}

fn named_format(code: &str) -> Option<Named> {
    Some(match code {
        "general number" => Named::General,
        "currency" => Named::Code("$#,##0.00;-$#,##0.00"),
        "fixed" => Named::Code("0.00"),
        "standard" => Named::Code("#,##0.00"),
        "percent" => Named::Code("0.00%"),
        "scientific" => Named::Code("0.00E+00"),
        "yes/no" => Named::Words("Yes", "No"),
        "true/false" => Named::Words("True", "False"),
        "on/off" => Named::Words("On", "Off"),
        "general date" => Named::Code("m/d/yyyy h:mm:ss AM/PM"),
        "long date" => Named::Code("dddd, mmmm d, yyyy"),
        "medium date" => Named::Code("dd-mmm-yy"),
        "short date" => Named::Code("m/d/yyyy"),
        "long time" => Named::Code("h:mm:ss AM/PM"),
        "medium time" => Named::Code("hh:mm AM/PM"),
        "short time" => Named::Code("hh:mm"),
        _ => return None,
    })
}

/// VBA writes minutes as `nn`; the spreadsheet format says `mm` (and tells them from months by
/// context).
fn minutes_as_spreadsheet(code: &str) -> String {
    let mut out = String::with_capacity(code.len());
    let mut in_quote = false;
    let chars: Vec<char> = code.chars().collect();
    let mut i = 0;
    while let Some(&c) = chars.get(i) {
        if c == '"' {
            in_quote = !in_quote;
        }
        if !in_quote && (c == 'n' || c == 'N') && matches!(chars.get(i + 1), Some('n' | 'N')) {
            out.push_str("mm");
            i += 2;
            continue;
        }
        out.push(c);
        i += 1;
    }
    out
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn fixed_decimals_and_padding() {
        assert_eq!(format(&Value::Double(1.234_567_89), "0.000000").unwrap(), "1.234568");
        assert_eq!(format(&Value::Double(2.0), "00").unwrap(), "02");
        assert_eq!(format(&Value::Double(12.5), "0").unwrap(), "13");
        assert_eq!(format(&Value::Double(-0.5), "0.0").unwrap(), "-0.5");
        assert_eq!(format(&Value::Double(1234.5), "#,##0.00").unwrap(), "1,234.50");
    }

    #[test]
    fn named_formats_and_dates() {
        assert_eq!(format(&Value::Double(0.256), "Percent").unwrap(), "25.60%");
        assert_eq!(format(&Value::Double(1.0), "Yes/No").unwrap(), "Yes");
        assert_eq!(format(&Value::Date(43831.0), "yyyy-mm-dd").unwrap(), "2020-01-01");
        assert_eq!(format(&Value::Date(43831.5), "hh:nn").unwrap(), "12:00");
    }
}
