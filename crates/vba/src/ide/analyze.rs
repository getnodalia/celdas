// SPDX-License-Identifier: MIT OR Apache-2.0
//! The checks that need to know which names exist: a name used that nothing declares, a variable
//! declared and never used. Together with the structure scan and the parser's own complaint they are
//! the marks an editor draws on a module.

use std::collections::{HashMap, HashSet};

use super::catalog::Catalog;
use super::lines::{Logical, Word, next_chars, prev_char, statements, words};
use super::scan::{Diagnostic, Scan, Severity, scan};
use crate::parser::parse_module;

/// Words of the language: they are never names to look up.
const KEYWORDS: &[&str] = &[
    "access", "alias", "and", "any", "append", "as", "base", "binary", "byref", "byval", "call", "case", "close", "compare", "const", "declare", "dim", "do", "each", "else",
    "elseif", "empty", "end", "enum", "eqv", "erase", "error", "event", "exit", "explicit", "false", "for", "friend", "function", "get", "global", "gosub", "goto", "if", "imp",
    "implements", "in", "input", "is", "let", "lib", "like", "line", "lock", "loop", "lset", "me", "mod", "new", "next", "not", "nothing", "null", "on", "open", "option",
    "optional", "or", "output", "paramarray", "preserve", "print", "private", "property", "ptrsafe", "public", "put", "raiseevent", "random", "read", "redim", "rem", "resume",
    "return", "rset", "seek", "select", "set", "shared", "static", "step", "stop", "sub", "tab", "spc", "text", "then", "to", "true", "type", "typeof", "until", "wend",
    "while", "width", "with", "withevents", "write", "xor", "attribute", "defbool", "defbyte", "defint", "deflng", "defsng", "defdbl", "defcur", "defdate", "defstr", "defobj",
    "defvar",
];

/// What the checks found in a module.
#[derive(Debug, Default)]
pub struct Analysis {
    pub scan: Scan,
    /// Every mark, in line order: the structure's, the parser's, the names'.
    pub diagnostics: Vec<Diagnostic>,
}

/// Reads `src` and everything wrong with it. `module` is its name, and `extra` lists names that exist
/// but that the catalog cannot know (other modules' procedures and variables, form controls).
pub fn analyze(module: &str, src: &str, catalog: &Catalog, extra: &[String]) -> Analysis {
    let scan = scan(src);
    let mut diagnostics = scan.diagnostics.clone();
    if !diagnostics.iter().any(|d| d.severity == Severity::Error)
        && let Err(e) = parse_module(module, src)
    {
        let line = e.line.max(1);
        let (start, end) = scan.lines.iter().find(|l| l.first <= line && line <= l.last).map_or((0, 1), |l| (l.indent, l.first_line_end()));
        diagnostics.push(Diagnostic::new(line, start, end, Severity::Error, "Syntax error: {message}", vec![("message", e.message)]));
    }
    names(&scan, catalog, extra, &mut diagnostics);
    diagnostics.sort_by_key(|d| (d.line, d.start));
    Analysis { scan, diagnostics }
}

/// The words of one statement that are names being looked up.
pub(super) fn used_names(stmt: &[char]) -> Vec<Word> {
    let ws = words(stmt);
    let lower: Vec<String> = ws.iter().map(|w| w.text.to_lowercase()).collect();
    if matches!(lower.first().map(String::as_str), Some("declare" | "attribute" | "option" | "implements" | "event")) {
        return Vec::new();
    }
    let mut out = Vec::new();
    for (i, w) in ws.iter().enumerate() {
        let l = &lower[i];
        if KEYWORDS.contains(&l.as_str()) {
            continue;
        }
        if prev_char(stmt, w.start) == Some('.') || next_chars(stmt, w.end, 2) == ":=" {
            continue;
        }
        // The type after `As` or `As New`, and the label after `GoTo`.
        let before = i.checked_sub(1).and_then(|p| lower.get(p)).map(String::as_str);
        let before2 = i.checked_sub(2).and_then(|p| lower.get(p)).map(String::as_str);
        if before == Some("as") || (before == Some("new") && before2 == Some("as")) || matches!(before, Some("goto" | "gosub" | "resume")) {
            continue;
        }
        // A type name in `TypeOf x Is Type` and the second word of `Case Is`/`Line Input` are keywords already.
        if before == Some("is") && lower.first().map(String::as_str) == Some("typeof") {
            continue;
        }
        out.push(w.clone());
    }
    out
}

/// A statement of a line: its text and where in the line's code it starts.
fn line_statements(line: &Logical) -> Vec<(usize, Vec<char>)> {
    statements(&line.code)
}

fn names(scan: &Scan, catalog: &Catalog, extra: &[String], out: &mut Vec<Diagnostic>) {
    let known: HashSet<String> = extra.iter().map(|n| n.to_lowercase()).collect();
    let mut globals: HashSet<String> = HashSet::new();
    for p in &scan.procs {
        globals.insert(p.name.to_lowercase());
    }
    for v in scan.vars.iter().filter(|v| v.proc.is_none()) {
        globals.insert(v.name.to_lowercase());
    }
    for t in &scan.types {
        globals.insert(t.name.to_lowercase());
        if t.is_enum {
            for (f, _) in &t.fields {
                globals.insert(f.to_lowercase());
            }
        }
    }
    for (l, _, _) in &scan.labels {
        globals.insert(l.to_lowercase());
    }
    // Where each local is declared, and how often each name is used in each procedure.
    let declared_at: HashSet<(u32, usize)> = scan.vars.iter().map(|v| (v.line, v.col)).collect();
    let mut uses: HashMap<(Option<usize>, String), usize> = HashMap::new();
    let explicit = scan.option_explicit;
    for (idx, line) in scan.lines.iter().enumerate() {
        let proc = scan.line_proc.get(idx).copied().flatten();
        let Some(pi) = proc else { continue };
        if scan.procs.get(pi).is_some_and(|p| p.line == line.first || p.end_line == p.line) {
            continue;
        }
        for (offset, stmt) in line_statements(line) {
            for w in used_names(&stmt) {
                let (l, c) = line.position(offset + w.start);
                if declared_at.contains(&(l, c)) {
                    continue;
                }
                let key = w.text.to_lowercase();
                *uses.entry((Some(pi), key.clone())).or_insert(0) += 1;
                let local = scan.vars.iter().any(|v| v.proc == Some(pi) && v.name.eq_ignore_ascii_case(&key));
                let defined = local || globals.contains(&key) || known.contains(&key) || catalog.global(&key).is_some() || catalog.class(&key).is_some();
                if explicit && !defined {
                    out.push(Diagnostic::new(l, c, c + w.text.chars().count(), Severity::Warning, "“{name}” is not defined", vec![("name", w.text.clone())]));
                }
            }
        }
    }
    for v in scan.vars.iter().filter(|v| v.proc.is_some() && !v.param && !v.constant) {
        let used = uses.get(&(v.proc, v.name.to_lowercase())).copied().unwrap_or(0);
        if used == 0 {
            out.push(Diagnostic::new(v.line, v.col, v.col + v.name.chars().count(), Severity::Hint, "Variable “{name}” is never used", vec![("name", v.name.clone())]));
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn marks(src: &str) -> Vec<String> {
        analyze("Module1", src, &Catalog::vba(), &[]).diagnostics.into_iter().map(|d| format!("{}:{}", d.line, d.message)).collect()
    }

    #[test]
    fn an_undeclared_name_is_flagged_only_under_option_explicit() {
        let src = "Sub A()\n    x = 1\n    MsgBox x\nEnd Sub\n";
        assert!(marks(src).is_empty());
        let m = marks(&format!("Option Explicit\n{src}"));
        assert_eq!(m, ["3:“x” is not defined", "4:“x” is not defined"]);
    }

    #[test]
    fn declared_names_members_arguments_and_labels_are_not_flagged() {
        let src = "Option Explicit\nPrivate Total As Long\nEnum Color\n    Red = 1\nEnd Enum\nFunction Twice(ByVal n As Long) As Long\n    Dim i As Long, s As String\n    On Error GoTo Fail\n    For i = 1 To n: s = s & Left(\"ab\", 1): Next i\n    Twice = n * 2 + Red + Total\n    Debug.Print s, vbCrLf\n    Exit Function\nFail:\n    Err.Clear\nEnd Function\nSub B()\n    Dim d As New Collection\n    d.Add 1, Key:=\"k\"\n    With d\n        .Add 2\n    End With\n    Twice 3\nEnd Sub\n";
        assert!(marks(src).is_empty(), "{:?}", marks(src));
    }

    #[test]
    fn an_unused_variable_is_a_hint_and_a_used_one_is_not() {
        let m = marks("Sub A()\n    Dim a As Long, b As Long\n    a = 1\nEnd Sub\n");
        assert_eq!(m, ["2:Variable “b” is never used"]);
    }

    #[test]
    fn the_parsers_own_complaint_is_added_when_the_blocks_are_fine() {
        let m = marks("Sub A()\n    x = = 1\nEnd Sub\n");
        assert_eq!(m.len(), 1, "{m:?}");
        assert!(m.first().is_some_and(|s| s.starts_with("2:Syntax error")), "{m:?}");
    }

    #[test]
    fn extra_names_are_known() {
        let a = analyze("M", "Option Explicit\nSub A()\n    Other\nEnd Sub\n", &Catalog::vba(), &["Other".to_string()]);
        assert!(a.diagnostics.is_empty(), "{:?}", a.diagnostics);
    }
}
