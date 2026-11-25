//! What the VBA IDE works out from text, without drawing: lines and columns, where a breakpoint
//! goes when lines are added above it, what accepting a suggestion does to the text, the search.
//! Pure functions, tested below.

use std::collections::BTreeSet;
use std::ops::Range;

/// The text the editor shows: lines end with `\n` (the project file keeps `\r\n`; the engine puts
/// them back on save).
pub fn normalize(code: &str) -> String {
    code.replace("\r\n", "\n").replace('\r', "\n")
}

/// Index of the char at byte `b` of `text` (the end of the text when `b` is past it).
pub fn char_of_byte(text: &str, b: usize) -> usize {
    text.char_indices().take_while(|(i, _)| *i < b).count()
}

/// Byte of char `c` of `text` (the end of the text when it is past it).
pub fn byte_of_char(text: &str, c: usize) -> usize {
    text.char_indices().nth(c).map_or(text.len(), |(b, _)| b)
}

/// The 1-based line and the column (chars) of char `c`.
pub fn line_col(text: &str, c: usize) -> (u32, usize) {
    let mut line = 1u32;
    let mut col = 0usize;
    for ch in text.chars().take(c) {
        if ch == '\n' {
            line = line.saturating_add(1);
            col = 0;
        } else {
            col += 1;
        }
    }
    (line, col)
}

/// The char index of `line` (1-based) and `col` (chars, clamped to the line).
pub fn char_at(text: &str, line: u32, col: usize) -> usize {
    let mut start = 0usize;
    let mut n = 1u32;
    let mut chars = text.chars().enumerate();
    while n < line {
        match chars.find(|(_, c)| *c == '\n') {
            Some((i, _)) => {
                start = i + 1;
                n += 1;
            }
            None => return text.chars().count(),
        }
    }
    let len = text.chars().skip(start).take_while(|c| *c != '\n').count();
    start + col.min(len)
}

/// The byte range of `line` (1-based) `start..end` columns, clamped to the line.
pub fn byte_range(text: &str, line: u32, start: usize, end: usize) -> Range<usize> {
    let a = byte_of_char(text, char_at(text, line, start));
    let b = byte_of_char(text, char_at(text, line, end.max(start)));
    a..b.max(a)
}

/// The text of `line` (1-based), without its newline.
pub fn line_text(text: &str, line: u32) -> &str {
    text.split('\n').nth((line as usize).saturating_sub(1)).unwrap_or("")
}

/// How many lines the text has (an empty text has one).
pub fn line_count(text: &str) -> u32 {
    u32::try_from(text.split('\n').count()).unwrap_or(u32::MAX)
}

/// Moves the lines in `marks` (breakpoints) so each stays on the statement it was on after the
/// text changed from `old` to `new`: lines below the edit shift by the lines it added or removed,
/// and a line the edit removed loses its mark.
pub fn shift_lines(old: &str, new: &str, marks: &BTreeSet<u32>) -> BTreeSet<u32> {
    let a: Vec<&str> = old.split('\n').collect();
    let b: Vec<&str> = new.split('\n').collect();
    let prefix = a.iter().zip(&b).take_while(|(x, y)| x == y).count();
    let suffix = a.iter().rev().zip(b.iter().rev()).take(a.len().min(b.len()) - prefix).take_while(|(x, y)| x == y).count();
    let old_end = a.len() - suffix; // first old line of the common tail (0-based)
    let new_end = b.len() - suffix;
    marks
        .iter()
        .filter_map(|&line| {
            let i = (line as usize).checked_sub(1)?;
            if i < prefix {
                Some(line)
            } else if i >= old_end {
                let moved = i + new_end - old_end;
                u32::try_from(moved + 1).ok()
            } else if a.len() == b.len() {
                // A line changed in place keeps its mark.
                Some(line)
            } else {
                None
            }
        })
        .collect()
}

/// What accepting a suggestion does: `from` is the column of the word being replaced, `caret` the
/// caret's char index. `insert` may hold `$0`, where the caret goes after. Returns the new text and
/// the caret. The indentation of the current line carries to the lines a snippet adds.
pub fn accept(text: &str, from: usize, caret: usize, insert: &str) -> (String, usize) {
    let (line, _) = line_col(text, caret);
    let start = char_at(text, line, from).min(caret);
    let indent: String = line_text(text, line).chars().take_while(|c| *c == ' ' || *c == '\t').collect();
    let mut body = String::new();
    for (i, part) in insert.split('\n').enumerate() {
        if i > 0 {
            body.push('\n');
            body.push_str(&indent);
        }
        body.push_str(part);
    }
    let (before, after) = match body.split_once("$0") {
        Some((a, b)) => (a.to_string(), b.to_string()),
        None => (body.clone(), String::new()),
    };
    let mut out = String::with_capacity(text.len() + body.len());
    out.push_str(text.get(..byte_of_char(text, start)).unwrap_or(""));
    out.push_str(&before);
    out.push_str(&after);
    out.push_str(text.get(byte_of_char(text, caret)..).unwrap_or(""));
    (out, start + before.chars().count())
}

/// The word (letters, digits, `_`) the char `c` is in or touches: its char range and its text.
pub fn word_at(text: &str, c: usize) -> Option<(Range<usize>, String)> {
    let chars: Vec<char> = text.chars().collect();
    let is_word = |ch: char| ch.is_alphanumeric() || ch == '_';
    let mut a = c.min(chars.len());
    let mut b = a;
    while a > 0 && chars.get(a - 1).is_some_and(|ch| is_word(*ch)) {
        a -= 1;
    }
    while chars.get(b).is_some_and(|ch| is_word(*ch)) {
        b += 1;
    }
    (a < b).then(|| (a..b, chars[a..b].iter().collect()))
}

/// How a search matches.
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
pub struct Match {
    pub case: bool,
    pub whole: bool,
}

/// Every place `needle` is in `text`, as byte ranges, in order and not overlapping.
pub fn find_all(text: &str, needle: &str, how: Match) -> Vec<Range<usize>> {
    if needle.is_empty() {
        return Vec::new();
    }
    // Lower-casing can change a string's length (some letters), so the search runs on chars.
    let hay: Vec<(usize, char)> = text.char_indices().collect();
    let pat: Vec<char> = needle.chars().collect();
    let same = |a: char, b: char| if how.case { a == b } else { a.to_lowercase().eq(b.to_lowercase()) };
    let is_word = |c: char| c.is_alphanumeric() || c == '_';
    let mut out = Vec::new();
    let mut i = 0;
    while i + pat.len() <= hay.len() {
        if hay[i..i + pat.len()].iter().zip(&pat).all(|((_, a), b)| same(*a, *b)) {
            let before_ok = i == 0 || !is_word(hay[i - 1].1);
            let after = hay.get(i + pat.len()).map(|(_, c)| *c);
            let after_ok = after.is_none_or(|c| !is_word(c));
            if !how.whole || (before_ok && after_ok) {
                let start = hay[i].0;
                let end = hay.get(i + pat.len()).map_or(text.len(), |(b, _)| *b);
                out.push(start..end);
                i += pat.len();
                continue;
            }
        }
        i += 1;
    }
    out
}

/// `text` with every match of `needle` replaced.
pub fn replace_all(text: &str, needle: &str, with: &str, how: Match) -> (String, usize) {
    let hits = find_all(text, needle, how);
    let mut out = String::with_capacity(text.len());
    let mut at = 0;
    for r in &hits {
        out.push_str(text.get(at..r.start).unwrap_or(""));
        out.push_str(with);
        at = r.end;
    }
    out.push_str(text.get(at..).unwrap_or(""));
    (out, hits.len())
}

/// Whether `name` is a name VBA accepts for a module.
pub fn valid_module_name(name: &str) -> bool {
    let name = name.trim();
    let mut chars = name.chars();
    chars.next().is_some_and(|c| c.is_alphabetic()) && chars.all(|c| c.is_alphanumeric() || c == '_') && name.chars().count() <= 31
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn lines_and_columns_round_trip() {
        let t = "Sub A()\n    x = 1\nEnd Sub\n";
        assert_eq!(line_col(t, 0), (1, 0));
        assert_eq!(line_col(t, 12), (2, 4));
        assert_eq!(char_at(t, 2, 4), 12);
        assert_eq!(char_at(t, 2, 99), 17, "a column past the line ends at the line");
        assert_eq!(char_at(t, 99, 0), t.chars().count());
        assert_eq!(line_text(t, 2), "    x = 1");
        assert_eq!(line_count(t), 4);
        assert_eq!(byte_range("ñandú x", 1, 6, 7), 8..9);
    }

    #[test]
    fn breakpoints_follow_their_statement() {
        let old = "a\nb\nc\nd";
        let marks: BTreeSet<u32> = [2, 4].into();
        // A line inserted above shifts both.
        assert_eq!(shift_lines(old, "x\na\nb\nc\nd", &marks), [3, 5].into());
        // A line removed above shifts them up.
        assert_eq!(shift_lines(old, "b\nc\nd", &marks), [1, 3].into());
        // Removing the marked line takes its mark.
        assert_eq!(shift_lines(old, "a\nc\nd", &marks), [3].into());
        // Typing within a line keeps it.
        assert_eq!(shift_lines(old, "a\nbb\nc\nd", &marks), [2, 4].into());
        // Nothing changed.
        assert_eq!(shift_lines(old, old, &marks), marks);
    }

    #[test]
    fn accepting_a_suggestion_replaces_the_typed_word() {
        assert_eq!(accept("    Dim x As Lo", 15 - 2, 15, "Long"), ("    Dim x As Long".to_string(), 17));
        // A snippet keeps the line's indentation on its new lines and puts the caret at `$0`.
        let (t, c) = accept("    If", 4, 6, "If $0 Then\nEnd If");
        assert_eq!(t, "    If  Then\n    End If");
        assert_eq!(c, 7);
        // Text after the caret stays.
        assert_eq!(accept("x = Ra + 1", 4, 6, "Range").0, "x = Range + 1");
    }

    #[test]
    fn words_and_searches() {
        assert_eq!(word_at("x = Total + 1", 6), Some((4..9, "Total".into())));
        assert_eq!(word_at("x = Total + 1", 9), Some((4..9, "Total".into())));
        assert_eq!(word_at("a  b", 2), None);
        let t = "Dim total, Total2, TOTAL";
        assert_eq!(find_all(t, "total", Match::default()).len(), 3);
        assert_eq!(find_all(t, "total", Match { whole: true, case: false }), vec![4..9, 19..24]);
        assert_eq!(find_all(t, "total", Match { whole: false, case: true }).len(), 1);
        assert_eq!(replace_all(t, "total", "n", Match { whole: true, case: false }), ("Dim n, Total2, n".to_string(), 2));
        assert!(find_all("abc", "", Match::default()).is_empty());
    }

    #[test]
    fn module_names() {
        assert!(valid_module_name("Module1"));
        assert!(valid_module_name("Cálculos_2"));
        assert!(!valid_module_name("1Module"));
        assert!(!valid_module_name("a b"));
        assert!(!valid_module_name(""));
    }

    #[test]
    fn crlf_becomes_lf() {
        assert_eq!(normalize("a\r\nb\rc\n"), "a\nb\nc\n");
    }
}
