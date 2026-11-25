// SPDX-License-Identifier: MIT OR Apache-2.0
//! Reading source the way an editor needs it: tolerant of everything. A module that does not
//! parse still has lines, statements and words, so the outline, the completion and the
//! diagnostics keep working while the person types.

use crate::lexer::physical_lines;

/// A statement line: the physical lines it spans, joined at `_` continuations.
#[derive(Clone, Debug, PartialEq)]
pub struct Logical {
    /// The first and last physical line (1-based).
    pub first: u32,
    pub last: u32,
    /// The code with the contents of strings and the comment blanked out; each physical line
    /// keeps its columns, and lines are joined by one space (see [`Logical::position`]).
    pub code: Vec<char>,
    /// Where each physical line starts in `code`.
    pub starts: Vec<usize>,
    /// The comment that ends the last physical line, without the apostrophe.
    pub comment: Option<String>,
    /// Characters of whitespace before the code on the first physical line.
    pub indent: usize,
    /// The first physical line is blank or only a comment.
    pub blank: bool,
}

impl Logical {
    /// The physical line and column of the character at `offset` in `code`.
    pub fn position(&self, offset: usize) -> (u32, usize) {
        let mut line = 0;
        for (i, start) in self.starts.iter().enumerate() {
            if *start <= offset {
                line = i;
            }
        }
        let start = self.starts.get(line).copied().unwrap_or(0);
        (self.first + line as u32, offset.saturating_sub(start))
    }

    /// The code as text.
    pub fn text(&self) -> String {
        self.code.iter().collect()
    }

    /// Where the code of the first physical line ends (columns), for marking a whole statement.
    pub fn first_line_end(&self) -> usize {
        let end = self.starts.get(1).map_or(self.code.len(), |s| s.saturating_sub(1));
        let first: Vec<char> = self.code.get(..end).unwrap_or(&[]).to_vec();
        first.iter().rposition(|c| !c.is_whitespace()).map_or(self.indent, |p| p + 1)
    }
}

/// A date or time literal such as `#1/2/2024#`: its length, so the text between the hashes is
/// not read as code. `Open f For Input As #1` is not one.
fn date_literal_len(chars: &[char], at: usize) -> Option<usize> {
    let first = *chars.get(at + 1)?;
    if !first.is_ascii_digit() {
        return None;
    }
    let mut j = at + 1;
    while let Some(&c) = chars.get(j) {
        if c == '#' {
            return Some(j - at + 1);
        }
        if !(c.is_ascii_digit() || matches!(c, '/' | '-' | ':' | '.' | ' ' | 'A' | 'P' | 'M' | 'a' | 'p' | 'm')) {
            return None;
        }
        j += 1;
    }
    None
}

/// `Rem` starting a comment: the word, at the start of a statement, followed by a space or the end.
fn rem_at(chars: &[char], at: usize, out: &[char]) -> bool {
    let word: String = chars.get(at..at + 3).map(|w| w.iter().collect()).unwrap_or_default();
    if !word.eq_ignore_ascii_case("rem") || chars.get(at + 3).is_some_and(|c| c.is_alphanumeric() || *c == '_') {
        return false;
    }
    match out.iter().rposition(|c| !c.is_whitespace()) {
        None => true,
        Some(p) => out.get(p) == Some(&':'),
    }
}

/// One physical line with strings blanked and the comment cut off.
pub fn mask_line(line: &str) -> (Vec<char>, Option<String>) {
    let chars: Vec<char> = line.chars().collect();
    let mut out: Vec<char> = Vec::with_capacity(chars.len());
    let mut i = 0;
    while let Some(&c) = chars.get(i) {
        match c {
            '"' => {
                out.push('"');
                i += 1;
                while let Some(&d) = chars.get(i) {
                    if d == '"' {
                        if chars.get(i + 1) == Some(&'"') {
                            out.extend([' ', ' ']);
                            i += 2;
                            continue;
                        }
                        out.push('"');
                        i += 1;
                        break;
                    }
                    out.push(' ');
                    i += 1;
                }
            }
            '\'' => return (out, Some(chars.get(i + 1..).unwrap_or(&[]).iter().collect())),
            '#' => match date_literal_len(&chars, i) {
                Some(n) => {
                    out.push('#');
                    out.extend(std::iter::repeat_n(' ', n - 2));
                    out.push('#');
                    i += n;
                }
                None => {
                    out.push('#');
                    i += 1;
                }
            },
            'R' | 'r' if rem_at(&chars, i, &out) => return (out, Some(chars.get(i + 3..).unwrap_or(&[]).iter().collect())),
            _ => {
                out.push(c);
                i += 1;
            }
        }
    }
    (out, None)
}

fn ends_with_continuation(code: &[char]) -> bool {
    let trimmed = code.iter().rposition(|c| !c.is_whitespace());
    match trimmed {
        Some(p) if code.get(p) == Some(&'_') => p > 0 && code.get(p - 1).is_some_and(|c| c.is_whitespace()),
        _ => false,
    }
}

/// The logical lines of `src`.
pub fn logical_lines(src: &str) -> Vec<Logical> {
    let physical = physical_lines(src);
    let mut out = Vec::new();
    let mut i = 0;
    while let Some(first) = physical.get(i) {
        let first_no = (i + 1) as u32;
        let indent = first.chars().take_while(|c| c.is_whitespace()).count();
        let (mut code, mut comment) = mask_line(first);
        let mut starts = vec![0usize];
        let blank = code.iter().all(|c| c.is_whitespace());
        let mut last = first_no;
        while ends_with_continuation(&code) && comment.is_none() {
            // The continuation mark becomes a space, so columns stay the physical ones.
            if let Some(p) = code.iter().rposition(|c| !c.is_whitespace()) {
                code[p] = ' ';
            }
            i += 1;
            let Some(next) = physical.get(i) else { break };
            let (more, c) = mask_line(next);
            code.push(' ');
            starts.push(code.len());
            code.extend(more);
            comment = c;
            last += 1;
        }
        i += 1;
        out.push(Logical { first: first_no, last, code, starts, comment, indent, blank });
    }
    out
}

/// Splits masked code into statements at `:` (not `:=`), with the offset each starts at.
pub fn statements(code: &[char]) -> Vec<(usize, Vec<char>)> {
    let mut out = Vec::new();
    let mut start = 0;
    let mut depth = 0i32;
    for (i, &c) in code.iter().enumerate() {
        match c {
            '(' => depth += 1,
            ')' => depth -= 1,
            ':' if depth <= 0 && code.get(i + 1) != Some(&'=') && !(i > 0 && code.get(i - 1) == Some(&':')) => {
                out.push((start, code.get(start..i).unwrap_or(&[]).to_vec()));
                start = i + 1;
            }
            _ => {}
        }
    }
    out.push((start, code.get(start..).unwrap_or(&[]).to_vec()));
    out.retain(|(_, s)| s.iter().any(|c| !c.is_whitespace()));
    out
}

/// A word of a statement: its text and where it starts (characters).
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Word {
    pub text: String,
    pub start: usize,
    pub end: usize,
}

pub fn is_word_start(c: char) -> bool {
    c.is_alphabetic() || c == '_'
}

pub fn is_word_char(c: char) -> bool {
    c.is_alphanumeric() || c == '_'
}

/// Identifier-like words of `code`, in order, with the character before and after each one.
pub fn words(code: &[char]) -> Vec<Word> {
    let mut out = Vec::new();
    let mut i = 0;
    while let Some(&c) = code.get(i) {
        if is_word_start(c) {
            let start = i;
            while code.get(i).is_some_and(|c| is_word_char(*c)) {
                i += 1;
            }
            out.push(Word { text: code.get(start..i).unwrap_or(&[]).iter().collect(), start, end: i });
        } else if c.is_ascii_digit() || (c == '&' && code.get(i + 1).is_some_and(|d| matches!(d, 'H' | 'h' | 'O' | 'o'))) {
            // A number (`1e5`, `&H1F`) is not a word.
            i += 1;
            while code.get(i).is_some_and(|c| c.is_alphanumeric() || *c == '.' || *c == '_') {
                i += 1;
            }
        } else {
            i += 1;
        }
    }
    out
}

/// The nearest non-blank character before `at`.
pub fn prev_char(code: &[char], at: usize) -> Option<char> {
    code.get(..at)?.iter().rev().find(|c| !c.is_whitespace()).copied()
}

/// The nearest non-blank characters from `at` on (up to `n`), as text.
pub fn next_chars(code: &[char], at: usize, n: usize) -> String {
    code.get(at..).unwrap_or(&[]).iter().filter(|c| !c.is_whitespace()).take(n).collect()
}

#[cfg(test)]
mod tests {
    use super::*;

    fn text(l: &Logical) -> String {
        l.text()
    }

    #[test]
    fn strings_and_comments_are_blanked_but_keep_their_columns() {
        let (code, comment) = mask_line("x = \"a'b\"\"c\" ' note");
        assert_eq!(code.iter().collect::<String>(), "x = \"      \" ");
        assert_eq!(comment.as_deref(), Some(" note"));
        let (code, comment) = mask_line("Rem a remark");
        assert!(code.is_empty());
        assert_eq!(comment.as_deref(), Some(" a remark"));
        // Rem inside a word is not a comment.
        assert_eq!(mask_line("Remainder = 1").1, None);
        assert_eq!(mask_line("x = 1: Rem done").1.as_deref(), Some(" done"));
    }

    #[test]
    fn a_date_literal_is_not_code_but_a_file_number_is() {
        let (code, _) = mask_line("d = #1/2/2024#");
        assert_eq!(code.iter().collect::<String>(), "d = #        #");
        let (code, _) = mask_line("Open f For Input As #1: Print #1, x");
        assert!(code.iter().collect::<String>().contains("Print #1, x"));
    }

    #[test]
    fn continued_lines_join_and_report_physical_positions() {
        let lines = logical_lines("x = 1 + _\n    2\ny = 3\n");
        assert_eq!(lines.len(), 2);
        let first = lines.first().unwrap();
        assert_eq!((first.first, first.last), (1, 2));
        assert_eq!(text(first).split_whitespace().collect::<Vec<_>>(), ["x", "=", "1", "+", "2"]);
        assert_eq!(first.position(0), (1, 0));
        let at_two = first.code.iter().position(|c| *c == '2').unwrap();
        assert_eq!(first.position(at_two), (2, 4));
        assert_eq!(lines.get(1).unwrap().first, 3);
    }

    #[test]
    fn statements_split_at_colons_but_not_named_arguments() {
        let code: Vec<char> = "a = 1: b = f(x:=2): c".chars().collect();
        let s = statements(&code);
        assert_eq!(s.len(), 3);
        assert_eq!(s.get(1).unwrap().1.iter().collect::<String>().trim(), "b = f(x:=2)");
    }

    #[test]
    fn words_skip_numbers_and_report_their_place() {
        let code: Vec<char> = "x1 = 2e5 + &H1F + abc".chars().collect();
        let w: Vec<String> = words(&code).into_iter().map(|w| w.text).collect();
        assert_eq!(w, ["x1", "abc"]);
    }
}
