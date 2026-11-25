// SPDX-License-Identifier: MIT OR Apache-2.0
//! Tidying a module: every line indented by how deeply its blocks are nested, the language's own
//! words spelled the usual way, and no trailing blanks. Names the person chose are never touched.

use super::lines::{Logical, logical_lines, prev_char, statements, words};
use super::scan::{Block, Cls, classify};

/// Spaces per level.
const INDENT: usize = 4;

/// The words that are spelled the same way everywhere, by their lower-case form.
const SPELLING: &[&str] = &[
    "Dim", "As", "ReDim", "Preserve", "Const", "Static", "Public", "Private", "Friend", "Global", "Sub", "Function", "Property", "End", "If", "Then", "Else", "ElseIf",
    "Select", "Case", "For", "To", "Step", "Next", "Each", "In", "Do", "While", "Until", "Loop", "Wend", "With", "Exit", "GoTo", "GoSub", "On", "Error", "Resume", "Set",
    "Let", "Call", "New", "Not", "And", "Or", "Xor", "Eqv", "Imp", "Mod", "Like", "Is", "TypeOf", "ByVal", "ByRef", "Optional", "ParamArray", "Option", "Explicit",
    "Declare", "Lib", "Type", "Enum", "True", "False", "Nothing", "Empty", "Null", "Me", "Boolean", "Byte", "Integer", "Long", "Single", "Double", "Currency", "Date",
    "String", "Variant", "Object", "Stop", "Erase", "WithEvents", "Implements", "Event", "RaiseEvent",
];

fn spelled(word: &str) -> Option<&'static str> {
    SPELLING.iter().find(|s| s.eq_ignore_ascii_case(word)).copied()
}

/// How a line is placed: the level it is printed at and the level the lines after it start at.
struct Level {
    depth: usize,
    /// For each open `Select Case`: whether a `Case` body is open.
    selects: Vec<bool>,
    stack: Vec<Block>,
}

impl Level {
    /// The level to print a statement of this class at; moves the level on for the lines after it.
    fn place(&mut self, cls: Cls) -> usize {
        match cls {
            Cls::Open(b) => {
                let at = self.depth;
                self.stack.push(b);
                self.depth += 1;
                if b == Block::Select {
                    self.selects.push(false);
                }
                at
            }
            Cls::Close(b, n) => {
                for _ in 0..n {
                    if let Some(at) = self.stack.iter().rposition(|x| *x == b) {
                        while self.stack.len() > at + 1 {
                            self.pop();
                        }
                        self.pop();
                    }
                }
                self.depth
            }
            Cls::ElseIf | Cls::Else => self.depth.saturating_sub(1),
            Cls::Case => {
                if self.selects.last().copied().unwrap_or(false) {
                    self.depth = self.depth.saturating_sub(1);
                }
                if let Some(open) = self.selects.last_mut() {
                    *open = true;
                }
                let at = self.depth;
                self.depth += 1;
                at
            }
            Cls::Other => self.depth,
        }
    }

    fn pop(&mut self) {
        if let Some(b) = self.stack.pop() {
            self.depth = self.depth.saturating_sub(1);
            if b == Block::Select && self.selects.pop().unwrap_or(false) {
                self.depth = self.depth.saturating_sub(1);
            }
        }
    }
}

/// `src`, tidied. Lines end with `\n`.
pub fn format(src: &str) -> String {
    let physical: Vec<&str> = crate::lexer::physical_lines(src);
    let logical = logical_lines(src);
    let mut out: Vec<String> = physical.iter().map(|l| l.trim_end().to_string()).collect();
    let mut level = Level { depth: 0, selects: Vec::new(), stack: Vec::new() };
    for line in &logical {
        let first_index = (line.first as usize).saturating_sub(1);
        if line.blank {
            // A comment keeps the level of the code around it.
            if let Some(text) = out.get_mut(first_index) {
                let body = text.trim_start().to_string();
                if !body.is_empty() {
                    *text = format!("{}{body}", " ".repeat(level.depth * INDENT));
                }
            }
            continue;
        }
        respell(line, &mut out);
        let stmts = statements(&line.code);
        let first_stmt = stmts.first().map(|(_, s)| s.as_slice()).unwrap_or(&[]);
        let cls = classify(first_stmt);
        let is_label = line.indent == 0 && is_label_line(&line.code);
        let at = if is_label { 0 } else { level.place(cls) };
        // Statements after a colon do not move the level (`For i = 1 To 2: x = i: Next i` is closed
        // on its own line); only the first one opens or closes.
        for (_, s) in stmts.iter().skip(1) {
            let c = classify(s);
            if matches!(c, Cls::Close(..) | Cls::Open(..)) {
                level.place(c);
            }
        }
        for (i, idx) in (line.first..=line.last).enumerate() {
            let Some(text) = out.get_mut((idx as usize).saturating_sub(1)) else { continue };
            let body = text.trim_start().to_string();
            let pad = if i == 0 { at * INDENT } else { at * INDENT + INDENT };
            *text = if body.is_empty() { String::new() } else { format!("{}{body}", " ".repeat(pad)) };
        }
    }
    let mut s = out.join("\n");
    if !s.is_empty() {
        s.push('\n');
    }
    s
}

/// `Done:` alone on a line.
fn is_label_line(code: &[char]) -> bool {
    let ws = words(code);
    let Some(w) = ws.first() else { return false };
    ws.len() == 1 && code.iter().rposition(|c| !c.is_whitespace()) == Some(w.end) && code.get(w.end) == Some(&':') && classify(&code[..w.end]) == Cls::Other
}

/// Writes the language's words in their usual spelling into the physical lines `line` covers.
fn respell(line: &Logical, out: &mut [String]) {
    for w in words(&line.code) {
        if prev_char(&line.code, w.start) == Some('.') {
            continue;
        }
        let Some(spelling) = spelled(&w.text) else { continue };
        if spelling == w.text || spelling.chars().count() != w.text.chars().count() {
            continue;
        }
        let (l, col) = line.position(w.start);
        let Some(text) = out.get_mut((l as usize).saturating_sub(1)) else { continue };
        let chars: Vec<char> = text.chars().collect();
        let original: String = chars.get(col..col + w.end - w.start).unwrap_or(&[]).iter().collect();
        if !original.eq_ignore_ascii_case(&w.text) {
            continue;
        }
        let mut next: Vec<char> = chars;
        for (k, c) in spelling.chars().enumerate() {
            if let Some(slot) = next.get_mut(col + k) {
                *slot = c;
            }
        }
        *text = next.into_iter().collect();
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn blocks_are_indented_and_keywords_spelled() {
        let src = "sub a()\ndim i as long\nfor i = 1 to 3\nif i = 2 then\nx = 1\nelse\nx = 2\nend if\nnext i\nend sub\n";
        let want = "Sub a()\n    Dim i As Long\n    For i = 1 To 3\n        If i = 2 Then\n            x = 1\n        Else\n            x = 2\n        End If\n    Next i\nEnd Sub\n";
        assert_eq!(format(src), want);
    }

    #[test]
    fn select_case_indents_its_cases_and_labels_stay_in_column_one() {
        let src = "Sub A()\nSelect Case x\nCase 1\ny = 1\nCase Else\ny = 2\nEnd Select\nOn Error GoTo Done\nDone:\nExit Sub\nEnd Sub\n";
        let want = "Sub A()\n    Select Case x\n        Case 1\n            y = 1\n        Case Else\n            y = 2\n    End Select\n    On Error GoTo Done\nDone:\n    Exit Sub\nEnd Sub\n";
        assert_eq!(format(src), want);
    }

    #[test]
    fn strings_comments_members_and_continuations_are_left_alone() {
        let src = "Sub A()\n  x = \"if then\" ' dim as\n  r.End = 1\n  y = 1 + _\n  2\nEnd Sub\n";
        let want = "Sub A()\n    x = \"if then\" ' dim as\n    r.End = 1\n    y = 1 + _\n        2\nEnd Sub\n";
        assert_eq!(format(src), want);
    }

    #[test]
    fn formatting_twice_changes_nothing_and_hostile_text_survives() {
        let src = "Sub A()\nIf a Then b = 1 Else b = 2\nFor i = 1 To 2: x = i: Next i\nWith r\n.A = 1\nEnd With\nEnd Sub\n";
        let once = format(src);
        assert_eq!(format(&once), once);
        for s in ["", "End If", "Next", "\u{0}", "Sub", "End Select\nEnd Select", "Case 1", "   \n\n  "] {
            let _ = format(s);
        }
    }

    #[test]
    fn a_single_line_if_does_not_indent_what_follows() {
        let out = format("Sub A()\nIf a Then b = 1\nx = 2\nEnd Sub\n");
        assert_eq!(out, "Sub A()\n    If a Then b = 1\n    x = 2\nEnd Sub\n");
    }
}
