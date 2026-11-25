// SPDX-License-Identifier: MIT OR Apache-2.0
//! What to offer while a person types: names that fit, the signature of the call being written, what
//! a name is, and where it was declared. Types are followed as far as the catalog and the
//! declarations allow (`ws.Range("A1").`, `With` blocks, `For Each`), and when they cannot be
//! followed nothing is guessed.

use super::catalog::{Catalog, Class, GlobalKind, Member, MemberKind};
use super::lines::{is_word_char, is_word_start, mask_line, statements};
use super::scan::{Scan, VarDecl};
use crate::lexer::physical_lines;

/// Most items one list offers.
const MAX_ITEMS: usize = 80;
/// Deepest chain of `With` and member lookups followed.
const MAX_CHAIN: usize = 8;

/// What an item is, for its icon.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum ItemKind {
    Keyword,
    Snippet,
    Function,
    Constant,
    Variable,
    Property,
    Method,
    Class,
    Procedure,
    Type,
}

/// One suggestion.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Item {
    pub label: String,
    pub kind: ItemKind,
    /// The signature or type, shown beside the label.
    pub detail: String,
    pub doc: String,
    /// What replaces the typed word. Lines end a snippet's block with `\n`; `$0` marks the caret.
    pub insert: String,
}

/// The suggestions at a point, and where the word they replace starts.
#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct Completion {
    /// Column (characters) of the first character of the typed word.
    pub from: usize,
    pub items: Vec<Item>,
}

/// The call being written.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Signature {
    pub label: String,
    /// Each parameter as written, in order.
    pub params: Vec<String>,
    /// Which parameter the caret is in.
    pub active: usize,
    pub doc: String,
}

/// What the caret is on or after.
struct Cursor {
    /// The masked code before the caret on its logical line.
    before: Vec<char>,
    /// The procedure the caret is in.
    proc: Option<usize>,
    /// The logical line (index into `scan.lines`).
    line: usize,
    /// The caret is inside a string or a comment.
    quiet: bool,
}

fn cursor(scan: &Scan, src: &str, line: u32, col: usize) -> Option<Cursor> {
    let idx = scan.lines.iter().position(|l| l.first <= line && line <= l.last)?;
    let l = scan.lines.get(idx)?;
    let start = l.starts.get((line - l.first) as usize).copied()?;
    let raw = physical_lines(src).get(line as usize - 1).copied().unwrap_or("");
    let prefix: String = raw.chars().take(col).collect();
    let (code, comment) = mask_line(&prefix);
    let quiet = comment.is_some() || code.iter().filter(|c| **c == '"').count() % 2 == 1;
    let end = (start + col).min(l.code.len());
    Some(Cursor { before: l.code.get(..end).unwrap_or(&[]).to_vec(), proc: scan.line_proc.get(idx).copied().flatten(), line: idx, quiet })
}

// ---------------------------------------------------------------- types

/// `Excel.Range` → `Range`, `New Collection` → `Collection`.
fn bare_type(ty: &str) -> &str {
    ty.rsplit('.').next().unwrap_or(ty)
}

/// One step of an expression: a name, and whether arguments followed it.
#[derive(Clone, Debug, PartialEq, Eq)]
struct Segment {
    name: String,
    called: bool,
}

/// The part of the code before `dot` that is the object (`a.b(1).c`), as text.
fn receiver(code: &[char], dot: usize) -> Vec<char> {
    let mut i = dot;
    loop {
        // Arguments right before: jump back over them.
        if i > 0 && code.get(i - 1) == Some(&')') {
            let mut depth = 0i32;
            let mut j = i;
            while j > 0 {
                j -= 1;
                match code.get(j) {
                    Some(')') => depth += 1,
                    Some('(') => {
                        depth -= 1;
                        if depth == 0 {
                            break;
                        }
                    }
                    _ => {}
                }
            }
            i = j;
        }
        let end = i;
        while i > 0 && code.get(i - 1).is_some_and(|c| is_word_char(*c) || matches!(c, '$' | '%' | '&' | '!' | '#' | '@')) {
            i -= 1;
        }
        if i == end && i > 0 && code.get(i - 1) != Some(&')') {
            // Nothing before the dot: a `With` member.
            return code.get(i..dot).unwrap_or(&[]).to_vec();
        }
        if i > 0 && code.get(i - 1) == Some(&'.') {
            i -= 1;
            continue;
        }
        break;
    }
    code.get(i..dot).unwrap_or(&[]).to_vec()
}

/// Splits `a.b(1).c` into its steps. A leading empty name means the receiver started with a dot.
fn segments(text: &[char]) -> Vec<Segment> {
    let mut out = Vec::new();
    let (mut name, mut called, mut depth) = (String::new(), false, 0i32);
    for &c in text {
        match c {
            '(' => {
                if depth == 0 {
                    called = true;
                }
                depth += 1;
            }
            ')' => depth -= 1,
            '.' if depth == 0 => {
                out.push(Segment { name: std::mem::take(&mut name), called });
                called = false;
            }
            _ if depth == 0 && !c.is_whitespace() => name.push(c),
            _ => {}
        }
    }
    out.push(Segment { name, called });
    out
}

/// A name written with a type character (`Name$`), without it.
fn plain(name: &str) -> &str {
    name.trim_end_matches(['$', '%', '&', '!', '#', '@'])
}

fn local<'a>(scan: &'a Scan, proc: Option<usize>, name: &str) -> Option<&'a VarDecl> {
    let name = plain(name);
    scan.vars.iter().filter(|v| v.proc == proc && proc.is_some()).chain(scan.vars.iter().filter(|v| v.proc.is_none())).find(|v| v.name.eq_ignore_ascii_case(name))
}

/// The class a first step names.
fn start_class(scan: &Scan, catalog: &Catalog, proc: Option<usize>, cursor_line: usize, seg: &Segment, depth: usize) -> Option<String> {
    if seg.name.is_empty() {
        return with_class(scan, catalog, proc, cursor_line, depth);
    }
    if seg.name.eq_ignore_ascii_case("me") {
        return None;
    }
    if let Some(v) = local(scan, proc, &seg.name) {
        let ty = v.ty.as_deref().map(bare_type)?;
        return step_item(catalog, ty, seg.called);
    }
    if let Some(p) = scan.procs.iter().find(|p| p.name.eq_ignore_ascii_case(&seg.name)) {
        return p.returns.as_deref().map(|t| bare_type(t).to_string());
    }
    if let Some(g) = catalog.global(plain(&seg.name)) {
        let ty = bare_type(&g.returns);
        return if ty.is_empty() { None } else { step_item(catalog, ty, seg.called && g.kind != GlobalKind::Function) };
    }
    catalog.class(&seg.name).map(|c| c.name.clone())
}

/// The class after applying arguments to something of class `ty`: an item of a collection.
fn step_item(catalog: &Catalog, ty: &str, called: bool) -> Option<String> {
    let class = catalog.class(ty)?;
    if called {
        return class.item.clone().filter(|i| catalog.class(i).is_some());
    }
    Some(class.name.clone())
}

fn member_class(catalog: &Catalog, class: &Class, seg: &Segment) -> Option<String> {
    let m = class.member(&seg.name)?;
    let ty = bare_type(&m.returns);
    if ty.is_empty() {
        return None;
    }
    let target = catalog.class(ty)?;
    if seg.called && m.params.is_empty() {
        return target.item.clone().filter(|i| catalog.class(i).is_some());
    }
    Some(target.name.clone())
}

/// The class of what `chain` names.
fn chain_class(scan: &Scan, catalog: &Catalog, proc: Option<usize>, cursor_line: usize, chain: &[Segment], depth: usize) -> Option<String> {
    if depth > MAX_CHAIN {
        return None;
    }
    let (first, rest) = chain.split_first()?;
    let mut class = start_class(scan, catalog, proc, cursor_line, first, depth + 1)?;
    for seg in rest {
        let c = catalog.class(&class)?;
        class = member_class(catalog, c, seg)?;
    }
    Some(class)
}

/// The class of the object of the `With` block that holds the line.
fn with_class(scan: &Scan, catalog: &Catalog, proc: Option<usize>, cursor_line: usize, depth: usize) -> Option<String> {
    let mut pending = 0usize;
    for idx in (0..=cursor_line).rev() {
        let line = scan.lines.get(idx)?;
        if scan.line_proc.get(idx).copied().flatten() != proc {
            return None;
        }
        for (_, stmt) in statements(&line.code).into_iter().rev() {
            let text: String = stmt.iter().collect();
            let t = text.trim().to_lowercase();
            if t.starts_with("end with") {
                pending += 1;
            } else if t.starts_with("with ") || t == "with" {
                if pending > 0 {
                    pending -= 1;
                    continue;
                }
                let body: Vec<char> = stmt.iter().skip_while(|c| c.is_whitespace()).copied().collect();
                let trimmed: Vec<char> = body.iter().skip(4).skip_while(|c| c.is_whitespace()).copied().collect();
                let chain = segments(&trimmed);
                return chain_class(scan, catalog, proc, idx.saturating_sub(1), &chain, depth + 1);
            }
        }
    }
    None
}

// ---------------------------------------------------------------- completion

fn rank_match(label: &str, prefix: &str) -> Option<u8> {
    if prefix.is_empty() {
        return Some(1);
    }
    let (l, p) = (label.to_lowercase(), prefix.to_lowercase());
    if l.starts_with(&p) {
        Some(0)
    } else if l.contains(&p) {
        Some(2)
    } else {
        None
    }
}

fn member_item(m: &Member) -> Item {
    Item {
        label: m.name.clone(),
        kind: if m.kind == MemberKind::Method { ItemKind::Method } else { ItemKind::Property },
        detail: m.signature(),
        doc: m.doc.clone(),
        insert: m.name.clone(),
    }
}

/// The typed word before the caret and its start column.
fn typed_word(before: &[char]) -> (usize, String) {
    let mut start = before.len();
    while start > 0 && before.get(start - 1).is_some_and(|c| is_word_char(*c)) {
        start -= 1;
    }
    let word: String = before.get(start..).unwrap_or(&[]).iter().collect();
    (start, word)
}

const TYPE_NAMES: &[&str] = &["Boolean", "Byte", "Currency", "Date", "Double", "Integer", "Long", "LongLong", "Object", "Single", "String", "Variant"];

const KEYWORDS: &[&str] = &[
    "Call", "Case", "Const", "Dim", "Do", "Each", "Else", "ElseIf", "End", "Erase", "Exit", "For", "Function", "GoTo", "If", "Loop", "Next", "Nothing", "On Error GoTo",
    "On Error Resume Next", "Option Explicit", "Private", "Public", "ReDim", "Select Case", "Set", "Static", "Step", "Sub", "Then", "To", "Until", "Wend", "While", "With",
    "True", "False", "Empty", "Null", "Me", "Not", "And", "Or", "Xor", "Mod", "Like", "Is", "New", "As", "ByVal", "ByRef", "Optional", "ParamArray", "Property",
];

/// (trigger, label, text). `$0` is where the caret ends up.
const SNIPPETS: &[(&str, &str, &str)] = &[
    ("if", "If … Then … End If", "If $0 Then\n    \nEnd If"),
    ("ife", "If … Else … End If", "If $0 Then\n    \nElse\n    \nEnd If"),
    ("for", "For … Next", "For i = 1 To $0\n    \nNext i"),
    ("foreach", "For Each … Next", "For Each item In $0\n    \nNext item"),
    ("do", "Do While … Loop", "Do While $0\n    \nLoop"),
    ("select", "Select Case", "Select Case $0\n    Case 1\n        \n    Case Else\n        \nEnd Select"),
    ("with", "With … End With", "With $0\n    \nEnd With"),
    ("sub", "Sub … End Sub", "Sub $0()\n    \nEnd Sub"),
    ("function", "Function … End Function", "Function $0() As Variant\n    \nEnd Function"),
    ("onerror", "On Error GoTo handler", "On Error GoTo Handler\n    $0\n    Exit Sub\nHandler:\n    MsgBox Err.Description"),
];

/// True when the text before the word ends with `As ` or `As New `.
fn after_as(before: &[char], word_start: usize) -> bool {
    let head: String = before.get(..word_start).unwrap_or(&[]).iter().collect();
    let h = head.trim_end().to_lowercase();
    h.ends_with(" as") || h.ends_with(" as new") || h == "as"
}

/// The suggestions at `line`/`col` (1-based line, character column) of `src`. `explicit` is a request
/// (Ctrl+Space): an empty word then lists everything, and otherwise nothing is offered for it.
pub fn complete(scan: &Scan, catalog: &Catalog, src: &str, line: u32, col: usize, explicit: bool) -> Completion {
    let Some(cur) = cursor(scan, src, line, col) else { return Completion::default() };
    if cur.quiet {
        return Completion::default();
    }
    let (start, word) = typed_word(&cur.before);
    let from = col.saturating_sub(word.chars().count());
    let dotted = start > 0 && cur.before.get(start - 1) == Some(&'.');
    let mut items = if dotted {
        member_items(scan, catalog, &cur, start - 1)
    } else if word.is_empty() && !explicit {
        Vec::new()
    } else if after_as(&cur.before, start) {
        type_items(scan, catalog)
    } else if word.chars().next().is_some_and(|c| !is_word_start(c)) {
        Vec::new()
    } else {
        name_items(scan, catalog, cur.proc)
    };
    items.retain(|i| rank_match(&i.label, &word).is_some());
    items.sort_by_key(|i| (rank_match(&i.label, &word).unwrap_or(9), order(i.kind), i.label.to_lowercase()));
    items.dedup_by(|a, b| a.label.eq_ignore_ascii_case(&b.label) && a.kind == b.kind);
    items.truncate(MAX_ITEMS);
    // A single suggestion that is the word itself says nothing.
    if items.len() == 1 && items.first().is_some_and(|i| i.label == word) {
        items.clear();
    }
    Completion { from, items }
}

fn order(kind: ItemKind) -> u8 {
    match kind {
        ItemKind::Variable => 0,
        ItemKind::Procedure => 1,
        ItemKind::Property | ItemKind::Method => 2,
        ItemKind::Function | ItemKind::Constant => 3,
        ItemKind::Class | ItemKind::Type => 4,
        ItemKind::Snippet => 5,
        ItemKind::Keyword => 6,
    }
}

fn member_items(scan: &Scan, catalog: &Catalog, cur: &Cursor, dot: usize) -> Vec<Item> {
    let text = receiver(&cur.before, dot);
    let chain = segments(&text);
    let Some(class) = chain_class(scan, catalog, cur.proc, cur.line, &chain, 0) else { return Vec::new() };
    catalog.class(&class).map(|c| c.members.iter().map(member_item).collect()).unwrap_or_default()
}

fn type_items(scan: &Scan, catalog: &Catalog) -> Vec<Item> {
    let mut out: Vec<Item> = TYPE_NAMES.iter().map(|t| Item { label: (*t).into(), kind: ItemKind::Type, detail: "type".into(), doc: String::new(), insert: (*t).into() }).collect();
    out.extend(catalog.classes.iter().map(|c| Item { label: c.name.clone(), kind: ItemKind::Class, detail: "class".into(), doc: c.doc.clone(), insert: c.name.clone() }));
    out.extend(scan.types.iter().map(|t| Item {
        label: t.name.clone(),
        kind: ItemKind::Type,
        detail: if t.is_enum { "Enum" } else { "Type" }.into(),
        doc: String::new(),
        insert: t.name.clone(),
    }));
    out
}

fn name_items(scan: &Scan, catalog: &Catalog, proc: Option<usize>) -> Vec<Item> {
    let mut out = Vec::new();
    for v in scan.visible_vars(proc) {
        out.push(Item {
            label: v.name.clone(),
            kind: if v.constant { ItemKind::Constant } else { ItemKind::Variable },
            detail: v.ty.clone().unwrap_or_default(),
            doc: String::new(),
            insert: v.name.clone(),
        });
    }
    for p in &scan.procs {
        out.push(Item { label: p.name.clone(), kind: ItemKind::Procedure, detail: p.signature(), doc: p.doc.clone(), insert: p.name.clone() });
    }
    for t in scan.types.iter().filter(|t| t.is_enum) {
        for (f, _) in &t.fields {
            out.push(Item { label: f.clone(), kind: ItemKind::Constant, detail: t.name.clone(), doc: String::new(), insert: f.clone() });
        }
    }
    for g in &catalog.globals {
        let kind = match g.kind {
            GlobalKind::Function => ItemKind::Function,
            GlobalKind::Constant => ItemKind::Constant,
            GlobalKind::Variable => ItemKind::Variable,
        };
        out.push(Item { label: g.name.clone(), kind, detail: g.signature(), doc: g.doc.clone(), insert: g.name.clone() });
    }
    for c in &catalog.classes {
        out.push(Item { label: c.name.clone(), kind: ItemKind::Class, detail: "class".into(), doc: c.doc.clone(), insert: c.name.clone() });
    }
    out.extend(KEYWORDS.iter().map(|k| Item { label: (*k).into(), kind: ItemKind::Keyword, detail: String::new(), doc: String::new(), insert: (*k).into() }));
    out.extend(SNIPPETS.iter().map(|(trigger, label, text)| Item {
        label: (*trigger).into(),
        kind: ItemKind::Snippet,
        detail: (*label).into(),
        doc: String::new(),
        insert: (*text).into(),
    }));
    out
}

// ---------------------------------------------------------------- signature help

/// The innermost call open at the caret: its callee text and the argument the caret is in.
fn open_call(before: &[char]) -> Option<(usize, usize, Option<String>)> {
    let mut depth = 0i32;
    let mut commas = vec![0usize];
    let mut i = before.len();
    while i > 0 {
        i -= 1;
        match before.get(i)? {
            ')' => {
                depth += 1;
                commas.push(0);
            }
            '(' => {
                if depth == 0 {
                    let named = named_argument(before, i + 1);
                    return Some((i, commas.last().copied().unwrap_or(0), named));
                }
                depth -= 1;
                commas.pop();
            }
            ',' => {
                if let Some(c) = commas.last_mut() {
                    *c += 1;
                }
            }
            _ => {}
        }
    }
    None
}

/// The name in `name:=` when the caret is in that argument.
fn named_argument(before: &[char], from: usize) -> Option<String> {
    let tail: String = before.get(from..).unwrap_or(&[]).iter().collect();
    let last = tail.rsplit(',').next().unwrap_or("");
    let (name, _) = last.split_once(":=")?;
    let name = name.trim();
    if name.is_empty() || !name.chars().all(is_word_char) { None } else { Some(name.to_string()) }
}

fn split_params(params: &str) -> Vec<String> {
    let mut out = Vec::new();
    let (mut depth, mut cur) = (0i32, String::new());
    for c in params.chars() {
        match c {
            '(' | '[' => depth += 1,
            ')' | ']' => depth -= 1,
            _ => {}
        }
        if c == ',' && depth == 0 {
            out.push(cur.trim().to_string());
            cur.clear();
        } else {
            cur.push(c);
        }
    }
    if !cur.trim().is_empty() {
        out.push(cur.trim().to_string());
    }
    out
}

/// The signature of the call the caret is inside, and the parameter it is in.
pub fn signature_help(scan: &Scan, catalog: &Catalog, src: &str, line: u32, col: usize) -> Option<Signature> {
    let cur = cursor(scan, src, line, col)?;
    if cur.quiet {
        return None;
    }
    let (open, mut active, named) = open_call(&cur.before)?;
    let callee_end = cur.before.get(..open)?.iter().rposition(|c| !c.is_whitespace())? + 1;
    let head = cur.before.get(..callee_end)?;
    let (start, name) = typed_word(head);
    if name.is_empty() {
        return None;
    }
    let (label, params, doc) = if start > 0 && head.get(start - 1) == Some(&'.') {
        let chain = segments(&receiver(head, start - 1));
        let class = chain_class(scan, catalog, cur.proc, cur.line, &chain, 0)?;
        let m = catalog.class(&class)?.member(&name)?;
        (m.signature(), split_params(&m.params), m.doc.clone())
    } else if let Some(p) = scan.procs.iter().find(|p| p.name.eq_ignore_ascii_case(&name)) {
        (p.signature(), p.params.iter().map(|q| q.text()).collect(), p.doc.clone())
    } else {
        let g = catalog.global(&name)?;
        (g.signature(), split_params(&g.params), g.doc.clone())
    };
    if let Some(n) = named
        && let Some(at) = params.iter().position(|p| p.trim_matches(['[', ']']).split_whitespace().next().is_some_and(|w| w.eq_ignore_ascii_case(&n)))
    {
        active = at;
    }
    Some(Signature { label, params, active, doc })
}

// ---------------------------------------------------------------- what a name is

/// The word under `col` and its columns.
pub fn word_at(scan: &Scan, line: u32, col: usize) -> Option<(String, usize, usize)> {
    let l = scan.lines.iter().find(|l| l.first <= line && line <= l.last)?;
    let start = l.starts.get((line - l.first) as usize).copied()?;
    let at = start + col;
    let code = &l.code;
    let on = |i: usize| code.get(i).is_some_and(|c| is_word_char(*c));
    let mut s = at;
    if !on(s) {
        if s > 0 && on(s - 1) { s -= 1 } else { return None }
    }
    let mut a = s;
    while a > 0 && on(a - 1) {
        a -= 1;
    }
    let mut b = s;
    while on(b) {
        b += 1;
    }
    let text: String = code.get(a..b)?.iter().collect();
    if text.chars().next().is_some_and(char::is_numeric) {
        return None;
    }
    let (_, ca) = l.position(a);
    Some((text, ca, ca + (b - a)))
}

/// A short description of the name under the caret, for a tooltip.
pub fn hover(scan: &Scan, catalog: &Catalog, src: &str, line: u32, col: usize) -> Option<(String, String)> {
    let (name, start, _) = word_at(scan, line, col)?;
    let cur = cursor(scan, src, line, start + 1)?;
    if cur.quiet {
        return None;
    }
    let before = cur.before.get(..cur.before.len().saturating_sub(1)).unwrap_or(&[]);
    let (_, _) = typed_word(before);
    let ws = before.len().saturating_sub(name.chars().count().saturating_sub(1));
    let dot = ws > 0 && before.get(ws - 1) == Some(&'.');
    if dot {
        let chain = segments(&receiver(before, ws - 1));
        let class = chain_class(scan, catalog, cur.proc, cur.line, &chain, 0)?;
        let m = catalog.class(&class)?.member(&name)?;
        return Some((m.signature(), m.doc.clone()));
    }
    if let Some(v) = local(scan, cur.proc, &name) {
        let kind = if v.param { "parameter" } else if v.constant { "constant" } else { "variable" };
        return Some((format!("{} As {}", v.name, v.ty.as_deref().unwrap_or("Variant")), kind.to_string()));
    }
    if let Some(p) = scan.procs.iter().find(|p| p.name.eq_ignore_ascii_case(&name)) {
        return Some((p.signature(), p.doc.clone()));
    }
    if let Some(g) = catalog.global(&name) {
        return Some((g.signature(), g.doc.clone()));
    }
    catalog.class(&name).map(|c| (c.name.clone(), c.doc.clone()))
}

/// Where the name under the caret is declared: the 1-based line and the column.
pub fn definition(scan: &Scan, line: u32, col: usize) -> Option<(u32, usize)> {
    let (name, _, _) = word_at(scan, line, col)?;
    let proc = scan.proc_at(line);
    if let Some(v) = scan.vars.iter().find(|v| v.proc == proc && v.proc.is_some() && v.name.eq_ignore_ascii_case(&name)) {
        return Some((v.line, v.col));
    }
    if let Some(p) = scan.procs.iter().find(|p| p.name.eq_ignore_ascii_case(&name)) {
        return Some((p.line, p.name_col));
    }
    if let Some(v) = scan.vars.iter().find(|v| v.proc.is_none() && v.name.eq_ignore_ascii_case(&name)) {
        return Some((v.line, v.col));
    }
    if let Some((_, l, _)) = scan.labels.iter().find(|(n, _, p)| n.eq_ignore_ascii_case(&name) && *p == proc) {
        return Some((*l, 0));
    }
    scan.types.iter().find(|t| t.name.eq_ignore_ascii_case(&name)).map(|t| (t.line, 0))
}

/// Every place `name` is written as a name in the module, as (line, start column, end column).
pub fn occurrences(scan: &Scan, name: &str) -> Vec<(u32, usize, usize)> {
    let mut out = Vec::new();
    for l in &scan.lines {
        let code = &l.code;
        let mut i = 0;
        while i < code.len() {
            if code.get(i).is_some_and(|c| is_word_start(*c)) && (i == 0 || !code.get(i - 1).is_some_and(|c| is_word_char(*c))) {
                let s = i;
                while code.get(i).is_some_and(|c| is_word_char(*c)) {
                    i += 1;
                }
                let w: String = code.get(s..i).unwrap_or(&[]).iter().collect();
                if w.eq_ignore_ascii_case(name) && super::lines::prev_char(code, s) != Some('.') {
                    let (ln, c) = l.position(s);
                    out.push((ln, c, c + (i - s)));
                }
            } else {
                i += 1;
            }
        }
    }
    out
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::ide::scan::scan;

    const EXCEL: &str = "class Range item=Range \"Cells\"\n  p Value : Variant \"The value\"\n  p Address([RowAbsolute]) : String ro \"Reference\"\n  p Cells : Range \"Cells\"\n  m Select\n  m Offset(RowOffset, ColumnOffset) : Range \"Moves\"\nclass Worksheet \"A sheet\"\n  p Name : String\n  m Range(Cell1, [Cell2]) : Range \"A range\"\nclass Sheets item=Worksheet \"Sheets\"\n  p Count : Long ro\nvar ActiveSheet : Worksheet \"The sheet in front\"\nvar Worksheets : Sheets \"All sheets\"\nfn Range(Cell1, [Cell2]) : Range \"A range\"\n";

    fn setup(src: &str) -> (Scan, Catalog) {
        let mut c = Catalog::vba();
        c.extend(EXCEL);
        (scan(src), c)
    }

    /// Completion at the end of the line holding `|`.
    fn at(src: &str) -> (String, Completion) {
        let marked = src.replace('|', "");
        let (line, col) = src.lines().enumerate().find_map(|(n, l)| l.find('|').map(|c| (n as u32 + 1, l[..c].chars().count()))).unwrap_or((1, 0));
        let (s, c) = setup(&marked);
        (marked.clone(), complete(&s, &c, &marked, line, col, false))
    }

    fn labels(c: &Completion) -> Vec<&str> {
        c.items.iter().map(|i| i.label.as_str()).collect()
    }

    #[test]
    fn members_follow_the_type_of_a_variable_a_call_and_a_with_block() {
        let (_, c) = at("Sub A()\n    Dim ws As Worksheet\n    ws.|\nEnd Sub\n");
        assert_eq!(labels(&c), ["Name", "Range"]);
        let (_, c) = at("Sub A()\n    Dim ws As Worksheet\n    ws.Range(\"A1\").Of|\nEnd Sub\n");
        assert_eq!(labels(&c), ["Offset"]);
        assert_eq!(c.from, "    ws.Range(\"A1\").".len());
        let (_, c) = at("Sub A()\n    With ActiveSheet.Range(\"A1\")\n        .|\n    End With\nEnd Sub\n");
        assert!(labels(&c).contains(&"Offset"), "{:?}", labels(&c));
        let (_, c) = at("Sub A()\n    Worksheets(1).|\nEnd Sub\n");
        assert_eq!(labels(&c), ["Name", "Range"]);
    }

    #[test]
    fn an_unknown_receiver_offers_nothing_and_strings_and_comments_are_quiet() {
        assert!(at("Sub A()\n    x.|\nEnd Sub\n").1.items.is_empty());
        assert!(at("Sub A()\n    s = \"Ms|\"\nEnd Sub\n").1.items.is_empty());
        assert!(at("Sub A()\n    ' Ms|\nEnd Sub\n").1.items.is_empty());
    }

    #[test]
    fn names_come_with_locals_first_then_the_library_and_keywords() {
        let (_, c) = at("Sub A()\n    Dim total As Long\n    Dim tags As String\n    t|\nEnd Sub\n");
        let l = labels(&c);
        assert_eq!(l.first(), Some(&"tags"), "{l:?}");
        assert!(l.contains(&"total") && l.contains(&"Time") && l.contains(&"Then"), "{l:?}");
        let (_, c) = at("Sub A()\n    Dim x As Wo|\nEnd Sub\n");
        assert_eq!(labels(&c), ["Worksheet"]);
    }

    #[test]
    fn nothing_is_offered_for_an_empty_word_unless_asked() {
        let src = "Sub A()\n    |\nEnd Sub\n";
        assert!(at(src).1.items.is_empty());
        let (s, c) = setup("Sub A()\n    \nEnd Sub\n");
        assert!(!complete(&s, &c, "Sub A()\n    \nEnd Sub\n", 2, 4, true).items.is_empty());
    }

    #[test]
    fn signature_help_names_the_call_and_its_active_parameter() {
        let src = "Sub A()\n    x = Mid(\"abc\", 2, |\nEnd Sub\n";
        let marked = src.replace('|', "");
        let (s, c) = setup(&marked);
        let sig = signature_help(&s, &c, &marked, 2, "    x = Mid(\"abc\", 2, ".chars().count()).unwrap();
        assert_eq!(sig.label, "Mid(String, Start, [Length]) As String");
        assert_eq!(sig.params, ["String", "Start", "[Length]"]);
        assert_eq!(sig.active, 2);
        let src = "Function Add(a As Long, b As Long) As Long\nEnd Function\nSub B()\n    Add 1, \nEnd Sub\n";
        let (s, c) = setup(src);
        let sig = signature_help(&s, &c, src, 4, "    Add 1, ".chars().count());
        assert!(sig.is_none(), "a call without parentheses has no tip yet");
        let src = "Function Add(a As Long, b As Long) As Long\nEnd Function\nSub B()\n    x = Add(1, \nEnd Sub\n";
        let (s, c) = setup(src);
        let sig = signature_help(&s, &c, src, 4, "    x = Add(1, ".chars().count()).unwrap();
        assert_eq!((sig.label.as_str(), sig.active), ("Function Add(a As Long, b As Long) As Long", 1));
    }

    #[test]
    fn hover_and_definition_find_the_declaration() {
        let src = "Sub A()\n    Dim total As Long\n    total = 1\nEnd Sub\n";
        let (s, c) = setup(src);
        let h = hover(&s, &c, src, 3, 6).unwrap();
        assert_eq!(h.0, "total As Long");
        assert_eq!(definition(&s, 3, 6), Some((2, 8)));
        assert_eq!(occurrences(&s, "TOTAL").len(), 2);
    }

    #[test]
    fn hostile_positions_do_not_panic() {
        let src = "Sub A(\n    x = (((\n\"";
        let (s, c) = setup(src);
        for line in 0..6 {
            for col in [0, 1, 5, 99] {
                let _ = complete(&s, &c, src, line, col, true);
                let _ = signature_help(&s, &c, src, line, col);
                let _ = hover(&s, &c, src, line, col);
                let _ = definition(&s, line, col);
            }
        }
    }
}
