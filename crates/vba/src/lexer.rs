//! VBA source to tokens. A module becomes one flat token stream with `Eol` between logical lines
//! (line continuations joined, comments dropped).

use std::rc::Rc;

use crate::value::Value;

/// A syntax error, with the line it was found on.
#[derive(Clone, Debug, PartialEq)]
pub struct SyntaxError {
    pub line: u32,
    pub message: String,
}

impl std::fmt::Display for SyntaxError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(f, "Line {}: {}", self.line, self.message)
    }
}

impl std::error::Error for SyntaxError {}

#[derive(Clone, Debug, PartialEq)]
pub enum Tok {
    /// An identifier or keyword, as written (type suffix removed), with that suffix.
    Ident(Rc<str>, Option<char>),
    Num(Value),
    Str(Rc<str>),
    Date(f64),
    /// Punctuation and operators: `+ - * / \ ^ & = < > <= >= <> ( ) , . ; : := ! #`.
    P(&'static str),
    Eol,
}

#[derive(Clone, Debug, PartialEq)]
pub struct Token {
    pub tok: Tok,
    pub line: u32,
    /// Whitespace came right before the token.
    pub space: bool,
}

/// The longest logical line or token stream the lexer accepts, so hostile source cannot exhaust
/// memory.
const MAX_TOKENS: usize = 4_000_000;

fn err(line: u32, message: impl Into<String>) -> SyntaxError {
    SyntaxError { line, message: message.into() }
}

/// Splits `src` into physical lines (any of `\r\n`, `\r`, `\n`).
pub(crate) fn physical_lines(src: &str) -> Vec<&str> {
    let mut out = Vec::new();
    let mut start = 0;
    let bytes = src.as_bytes();
    let mut i = 0;
    while i < bytes.len() {
        match bytes.get(i) {
            Some(b'\r') => {
                out.push(src.get(start..i).unwrap_or(""));
                if bytes.get(i + 1) == Some(&b'\n') {
                    i += 1;
                }
                start = i + 1;
            }
            Some(b'\n') => {
                out.push(src.get(start..i).unwrap_or(""));
                start = i + 1;
            }
            _ => {}
        }
        i += 1;
    }
    if start < src.len() {
        out.push(src.get(start..).unwrap_or(""));
    }
    out
}

/// Tokenizes a module's source.
pub fn lex(src: &str) -> Result<Vec<Token>, SyntaxError> {
    let lines = physical_lines(src);
    let mut out: Vec<Token> = Vec::new();
    let mut i = 0;
    while let Some(first) = lines.get(i) {
        let line_no = (i + 1) as u32;
        let mut logical = first.to_string();
        // A trailing " _" continues the statement on the next line.
        while logical.trim_end().ends_with(" _") || logical.trim_end().ends_with("\t_") {
            let t = logical.trim_end();
            logical = t.get(..t.len() - 1).unwrap_or("").to_string();
            i += 1;
            match lines.get(i) {
                Some(next) => logical.push_str(next),
                None => break,
            }
        }
        i += 1;
        lex_line(&logical, line_no, &mut out)?;
        if out.len() > MAX_TOKENS {
            return Err(err(line_no, "the module is too large"));
        }
    }
    Ok(out)
}

fn is_ident_start(c: char) -> bool {
    c.is_alphabetic() || c == '_'
}

fn is_ident_char(c: char) -> bool {
    c.is_alphanumeric() || c == '_'
}

fn lex_line(text: &str, line: u32, out: &mut Vec<Token>) -> Result<(), SyntaxError> {
    let chars: Vec<char> = text.chars().collect();
    let mut i = 0;
    let start_len = out.len();
    let mut space = false;
    while let Some(&c) = chars.get(i) {
        if c == ' ' || c == '\t' || c == '\u{a0}' || c == '\u{c}' {
            space = true;
            i += 1;
            continue;
        }
        if c == '\'' {
            break;
        }
        let at_statement_start = out.len() == start_len || matches!(out.last().map(|t| &t.tok), Some(Tok::P(":")));
        let tok = if c == '"' {
            lex_string(&chars, &mut i, line)?
        } else if c.is_ascii_digit() || (c == '.' && chars.get(i + 1).is_some_and(char::is_ascii_digit)) {
            lex_number(&chars, &mut i, line)?
        } else if c == '&' && matches!(chars.get(i + 1), Some('H' | 'h' | 'O' | 'o')) {
            lex_radix(&chars, &mut i, line)?
        } else if is_ident_start(c) {
            let (name, suffix) = lex_ident(&chars, &mut i);
            if at_statement_start && name.eq_ignore_ascii_case("rem") && chars.get(i).is_none_or(|c| c.is_whitespace()) {
                break;
            }
            Tok::Ident(Rc::from(name.as_str()), suffix)
        } else if c == '[' {
            // [Bracketed name]
            let mut j = i + 1;
            let mut name = String::new();
            while let Some(&ch) = chars.get(j) {
                if ch == ']' {
                    break;
                }
                name.push(ch);
                j += 1;
            }
            i = j + 1;
            Tok::Ident(Rc::from(name.as_str()), None)
        } else if c == '#' {
            match lex_date_or_hash(&chars, &mut i) {
                Some(t) => t,
                None => return Err(err(line, "invalid date literal")),
            }
        } else {
            lex_punct(&chars, &mut i, line)?
        };
        out.push(Token { tok, line, space });
        space = false;
    }
    if out.len() > start_len {
        out.push(Token { tok: Tok::Eol, line, space: false });
    }
    Ok(())
}

fn lex_string(chars: &[char], i: &mut usize, line: u32) -> Result<Tok, SyntaxError> {
    let mut s = String::new();
    *i += 1;
    loop {
        match chars.get(*i) {
            None => return Err(err(line, "string literal is not closed")),
            Some('"') => {
                if chars.get(*i + 1) == Some(&'"') {
                    s.push('"');
                    *i += 2;
                } else {
                    *i += 1;
                    break;
                }
            }
            Some(&ch) => {
                s.push(ch);
                *i += 1;
            }
        }
    }
    Ok(Tok::Str(Rc::from(s.as_str())))
}

fn lex_ident(chars: &[char], i: &mut usize) -> (String, Option<char>) {
    let mut name = String::new();
    while let Some(&ch) = chars.get(*i) {
        if is_ident_char(ch) {
            name.push(ch);
            *i += 1;
        } else {
            break;
        }
    }
    let mut suffix = None;
    if let Some(&s) = chars.get(*i)
        && matches!(s, '$' | '%' | '&' | '!' | '#' | '@')
        && !is_ident_char(chars.get(*i + 1).copied().unwrap_or(' '))
        // `x& y` is not a suffix when it is the concatenation operator followed by a name; a
        // suffix is glued to the name and followed by something that is not a name start.
        && !(s == '&' && chars.get(*i + 1).is_some_and(|c| *c == ' ' && chars.get(*i + 2).is_some_and(|d| is_ident_start(*d) || *d == '"')))
    {
        suffix = Some(s);
        *i += 1;
    }
    (name, suffix)
}

fn lex_number(chars: &[char], i: &mut usize, line: u32) -> Result<Tok, SyntaxError> {
    let mut text = String::new();
    let mut is_float = false;
    while let Some(&ch) = chars.get(*i) {
        if ch.is_ascii_digit() {
            text.push(ch);
        } else if ch == '.' && !is_float {
            is_float = true;
            text.push(ch);
        } else if (ch == 'e' || ch == 'E' || ch == 'd' || ch == 'D')
            && (chars.get(*i + 1).is_some_and(char::is_ascii_digit)
                || (matches!(chars.get(*i + 1), Some('+' | '-')) && chars.get(*i + 2).is_some_and(char::is_ascii_digit)))
        {
            is_float = true;
            text.push('e');
            if let Some(&sign) = chars.get(*i + 1)
                && (sign == '+' || sign == '-')
            {
                text.push(sign);
                *i += 1;
            }
        } else {
            break;
        }
        *i += 1;
    }
    let suffix = match chars.get(*i) {
        Some(&s @ ('%' | '&' | '!' | '#' | '@')) if !chars.get(*i + 1).is_some_and(|c| is_ident_char(*c)) => {
            *i += 1;
            Some(s)
        }
        _ => None,
    };
    let n: f64 = text.parse().map_err(|_| err(line, format!("invalid number `{text}`")))?;
    let v = match suffix {
        Some('%') => Value::Int(i16::try_from(n as i64).map_err(|_| err(line, "overflow"))?),
        Some('&') => Value::Long(i32::try_from(n as i64).map_err(|_| err(line, "overflow"))?),
        Some('!') => Value::Single(n as f32),
        Some('#') | Some('@') => Value::Double(n),
        _ if is_float => Value::Double(n),
        _ => {
            if n <= 32767.0 {
                Value::Int(n as i16)
            } else if n <= 2_147_483_647.0 {
                Value::Long(n as i32)
            } else {
                Value::Double(n)
            }
        }
    };
    Ok(Tok::Num(v))
}

fn lex_radix(chars: &[char], i: &mut usize, line: u32) -> Result<Tok, SyntaxError> {
    let radix = if matches!(chars.get(*i + 1), Some('H' | 'h')) { 16 } else { 8 };
    *i += 2;
    let mut digits = String::new();
    while let Some(&ch) = chars.get(*i) {
        if ch.is_digit(radix) {
            digits.push(ch);
            *i += 1;
        } else {
            break;
        }
    }
    let n = u64::from_str_radix(&digits, radix).map_err(|_| err(line, "invalid number"))?;
    let long_suffix = chars.get(*i) == Some(&'&');
    if long_suffix {
        *i += 1;
    }
    let v = if long_suffix || n > 0xFFFF {
        Value::Long(u32::try_from(n).map_err(|_| err(line, "overflow"))? as i32)
    } else {
        Value::Int(n as u16 as i16)
    };
    Ok(Tok::Num(v))
}

/// `#date#` or a lone `#` (a file number marker).
fn lex_date_or_hash(chars: &[char], i: &mut usize) -> Option<Tok> {
    let mut j = *i + 1;
    let mut body = String::new();
    while let Some(&ch) = chars.get(j) {
        if ch == '#' {
            if let Some(serial) = parse_date_literal(&body) {
                *i = j + 1;
                return Some(Tok::Date(serial));
            }
            break;
        }
        if !(ch.is_ascii_alphanumeric() || matches!(ch, '/' | ':' | '-' | ' ' | '.' | ',')) {
            break;
        }
        body.push(ch);
        j += 1;
    }
    *i += 1;
    Some(Tok::P("#"))
}

/// The serial of a date literal such as `1/31/2020`, `2020-01-31 10:30`, `10:30:15 PM`.
pub fn parse_date_literal(s: &str) -> Option<f64> {
    let s = s.trim();
    if s.is_empty() || !s.chars().any(|c| c.is_ascii_digit()) {
        return None;
    }
    let upper = s.to_ascii_uppercase();
    let (body, pm, am) = if let Some(b) = upper.strip_suffix("PM") {
        (b.trim().to_string(), true, false)
    } else if let Some(b) = upper.strip_suffix("AM") {
        (b.trim().to_string(), false, true)
    } else {
        (upper.clone(), false, false)
    };
    let mut date_part: Option<&str> = None;
    let mut time_part: Option<&str> = None;
    for piece in body.split_whitespace() {
        if piece.contains(':') {
            time_part = Some(piece);
        } else {
            date_part = Some(piece);
        }
    }
    let mut serial = 0.0;
    if let Some(d) = date_part {
        let nums: Vec<&str> = d.split(['/', '-', '.']).collect();
        let n: Vec<i64> = nums.iter().filter_map(|p| p.parse().ok()).collect();
        if n.len() != nums.len() {
            return None;
        }
        let (y, m, day) = match n.as_slice() {
            [a, b, c] if *a > 31 => (*a, *b, *c),
            [m, d, y] => (if *y < 100 { 2000 + *y } else { *y }, *m, *d),
            _ => return None,
        };
        serial = celdas_core::date::serial_from_ymd(celdas_core::DateSystem::D1900, y, m, day)?;
    }
    if let Some(t) = time_part {
        let parts: Vec<f64> = t.split(':').filter_map(|p| p.parse().ok()).collect();
        let (mut h, m, sec) = match parts.as_slice() {
            [h, m] => (*h, *m, 0.0),
            [h, m, s] => (*h, *m, *s),
            _ => return None,
        };
        if pm && h < 12.0 {
            h += 12.0;
        }
        if am && h == 12.0 {
            h = 0.0;
        }
        serial += celdas_core::date::time_fraction(h, m, sec);
    }
    Some(serial)
}

fn lex_punct(chars: &[char], i: &mut usize, line: u32) -> Result<Tok, SyntaxError> {
    let c = chars.get(*i).copied().unwrap_or(' ');
    let next = chars.get(*i + 1).copied();
    let (p, len): (&'static str, usize) = match (c, next) {
        ('<', Some('=')) => ("<=", 2),
        ('>', Some('=')) => (">=", 2),
        ('<', Some('>')) => ("<>", 2),
        ('=', Some('<')) => ("<=", 2),
        ('=', Some('>')) => (">=", 2),
        (':', Some('=')) => (":=", 2),
        ('+', _) => ("+", 1),
        ('-', _) => ("-", 1),
        ('*', _) => ("*", 1),
        ('/', _) => ("/", 1),
        ('\\', _) => ("\\", 1),
        ('^', _) => ("^", 1),
        ('&', _) => ("&", 1),
        ('=', _) => ("=", 1),
        ('<', _) => ("<", 1),
        ('>', _) => (">", 1),
        ('(', _) => ("(", 1),
        (')', _) => (")", 1),
        (',', _) => (",", 1),
        ('.', _) => (".", 1),
        (';', _) => (";", 1),
        (':', _) => (":", 1),
        ('!', _) => ("!", 1),
        ('?', _) => ("?", 1),
        _ => return Err(err(line, format!("unexpected character `{c}`"))),
    };
    *i += len;
    Ok(Tok::P(p))
}

#[cfg(test)]
mod tests {
    use super::*;

    fn toks(src: &str) -> Vec<Tok> {
        lex(src).unwrap().into_iter().map(|t| t.tok).collect()
    }

    #[test]
    fn continuation_comments_and_strings() {
        let t = toks("x = \"a\"\"b\" & _\n   \"c\" ' note\nRem skipped\ny = 1");
        assert_eq!(t.iter().filter(|t| **t == Tok::Eol).count(), 2);
        assert!(t.contains(&Tok::Str(Rc::from("a\"b"))));
        assert!(t.contains(&Tok::Str(Rc::from("c"))));
        assert!(!t.iter().any(|t| matches!(t, Tok::Ident(n, _) if n.eq_ignore_ascii_case("rem"))));
    }

    #[test]
    fn numbers_by_suffix() {
        assert_eq!(toks("2#")[0], Tok::Num(Value::Double(2.0)));
        assert_eq!(toks("1E-06")[0], Tok::Num(Value::Double(1e-6)));
        assert_eq!(toks("40000")[0], Tok::Num(Value::Long(40000)));
        assert_eq!(toks("&HFF")[0], Tok::Num(Value::Int(255)));
        assert_eq!(toks("&HFFFF")[0], Tok::Num(Value::Int(-1)));
        assert_eq!(toks("5!")[0], Tok::Num(Value::Single(5.0)));
    }

    #[test]
    fn hash_is_a_file_number_or_a_date() {
        let t = toks("Print #n, x");
        assert!(t.contains(&Tok::P("#")));
        let t = toks("d = #1/2/2020#");
        assert!(t.iter().any(|t| matches!(t, Tok::Date(_))));
    }

    #[test]
    fn suffixes_stay_off_names() {
        assert_eq!(toks("Left$(a, 2)")[0], Tok::Ident(Rc::from("Left"), Some('$')));
        assert_eq!(toks("a & b")[1], Tok::P("&"));
        assert_eq!(toks("a& b")[1], Tok::P("&"));
    }

    #[test]
    fn unterminated_string_is_an_error() {
        assert!(lex("x = \"abc").is_err());
    }
}
