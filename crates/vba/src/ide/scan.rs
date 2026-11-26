// SPDX-License-Identifier: MIT OR Apache-2.0
//! The structure of a module, read line by line and without a parser: its procedures, variables,
//! types and enumerations, and the mistakes in how its blocks open and close. It works on source
//! that does not compile, which is the source an editor is usually looking at.

use super::lines::{Logical, Word, is_word_char, logical_lines, statements, words};
use crate::ast::ProcKind;

/// How serious a mark is.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum Severity {
    Error,
    Warning,
    /// Faded, not underlined as wrong (an unused variable).
    Hint,
}

/// A mark on a stretch of one line.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Diagnostic {
    /// 1-based line.
    pub line: u32,
    /// Columns (characters) from `start` up to, not including, `end`.
    pub start: usize,
    pub end: usize,
    pub severity: Severity,
    /// The message with its values filled in.
    pub message: String,
    /// The message before its values were filled in (`{name}` stands for each), so a translated
    /// interface can translate it, and the values.
    pub template: &'static str,
    pub args: Vec<(&'static str, String)>,
}

impl Diagnostic {
    pub(super) fn new(
        line: u32,
        start: usize,
        end: usize,
        severity: Severity,
        template: &'static str,
        args: Vec<(&'static str, String)>,
    ) -> Diagnostic {
        let mut message = template.to_string();
        for (k, v) in &args {
            message = message.replace(&format!("{{{k}}}"), v);
        }
        Diagnostic { line, start, end: end.max(start + 1), severity, message, template, args }
    }
}

/// A parameter of a procedure.
#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct ParamDecl {
    pub name: String,
    pub ty: Option<String>,
    pub optional: bool,
    pub by_val: bool,
    pub param_array: bool,
    pub default: Option<String>,
}

impl ParamDecl {
    /// `[ByVal x As Long = 1]`, as a signature shows it.
    pub fn text(&self) -> String {
        let mut s = String::new();
        if self.param_array {
            s.push_str("ParamArray ");
        } else if self.by_val {
            s.push_str("ByVal ");
        }
        s.push_str(&self.name);
        if let Some(t) = &self.ty {
            s.push_str(" As ");
            s.push_str(t);
        }
        if let Some(d) = &self.default {
            s.push_str(" = ");
            s.push_str(d);
        }
        if self.optional { format!("[{s}]") } else { s }
    }
}

/// A Sub, Function or Property of the module.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct ProcDecl {
    pub name: String,
    pub kind: ProcKind,
    pub public: bool,
    /// The line of its header, and the line of its `End` (0 when it is never closed).
    pub line: u32,
    pub end_line: u32,
    /// Where its name starts on the header line.
    pub name_col: usize,
    pub params: Vec<ParamDecl>,
    pub returns: Option<String>,
    /// The comment lines right above the header.
    pub doc: String,
}

impl ProcDecl {
    /// `Function Name(a As Long) As String`.
    pub fn signature(&self) -> String {
        let kind = match self.kind {
            ProcKind::Sub => "Sub",
            ProcKind::Function => "Function",
            ProcKind::PropertyGet => "Property Get",
            ProcKind::PropertyLet => "Property Let",
            ProcKind::PropertySet => "Property Set",
        };
        let params: Vec<String> = self.params.iter().map(ParamDecl::text).collect();
        let mut s = format!("{kind} {}({})", self.name, params.join(", "));
        if let Some(r) = &self.returns {
            s.push_str(" As ");
            s.push_str(r);
        }
        s
    }
}

/// A variable or constant.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct VarDecl {
    pub name: String,
    /// The type as written (`Long`, `Excel.Range`), or from its suffix.
    pub ty: Option<String>,
    pub line: u32,
    pub col: usize,
    /// The procedure it belongs to (an index into [`Scan::procs`]), or `None` at module level.
    pub proc: Option<usize>,
    pub constant: bool,
    pub public: bool,
    pub param: bool,
}

/// A user-defined type or an enumeration.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct TypeDecl {
    pub name: String,
    pub is_enum: bool,
    pub line: u32,
    /// Fields (`Name As Type`) or members.
    pub fields: Vec<(String, Option<String>)>,
}

/// What a module holds, read without parsing it.
#[derive(Clone, Debug, Default)]
pub struct Scan {
    pub option_explicit: bool,
    pub procs: Vec<ProcDecl>,
    pub vars: Vec<VarDecl>,
    pub types: Vec<TypeDecl>,
    /// Labels (`Done:`) with their line.
    pub labels: Vec<(String, u32, Option<usize>)>,
    /// Mistakes in how blocks open and close, and a few more that need only the lines.
    pub diagnostics: Vec<Diagnostic>,
    pub lines: Vec<Logical>,
    /// The procedure each logical line (by index into `lines`) is in.
    pub line_proc: Vec<Option<usize>>,
}

impl Scan {
    /// The procedure that holds `line` (1-based).
    pub fn proc_at(&self, line: u32) -> Option<usize> {
        self.procs.iter().position(|p| line >= p.line && (p.end_line == 0 || line <= p.end_line))
    }

    /// The variables a procedure can see: its own, then the module's.
    pub fn visible_vars(&self, proc: Option<usize>) -> impl Iterator<Item = &VarDecl> {
        self.vars.iter().filter(move |v| v.proc.is_none() || v.proc == proc)
    }
}

// ---------------------------------------------------------------- classification

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(super) enum Block {
    Proc,
    If,
    For,
    Do,
    While,
    Select,
    With,
    Type,
    Enum,
}

impl Block {
    fn open_word(self) -> &'static str {
        match self {
            Block::Proc => "Sub",
            Block::If => "If",
            Block::For => "For",
            Block::Do => "Do",
            Block::While => "While",
            Block::Select => "Select Case",
            Block::With => "With",
            Block::Type => "Type",
            Block::Enum => "Enum",
        }
    }

    fn close_word(self, proc: Option<ProcKind>) -> &'static str {
        match self {
            Block::Proc => match proc {
                Some(ProcKind::Function) => "End Function",
                Some(ProcKind::Sub) | None => "End Sub",
                Some(_) => "End Property",
            },
            Block::If => "End If",
            Block::For => "Next",
            Block::Do => "Loop",
            Block::While => "Wend",
            Block::Select => "End Select",
            Block::With => "End With",
            Block::Type => "End Type",
            Block::Enum => "End Enum",
        }
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(super) enum Cls {
    Open(Block),
    /// `End X`, `Next`, `Loop`, `Wend`; for `Next i, j` the count is 2.
    Close(Block, usize),
    ElseIf,
    Else,
    Case,
    Other,
}

const MODIFIERS: &[&str] = &["public", "private", "friend", "static", "global"];

/// The index of the first word that is not a modifier.
fn skip_modifiers(w: &[String]) -> usize {
    let mut i = 0;
    while w.get(i).is_some_and(|x| MODIFIERS.contains(&x.as_str())) {
        i += 1;
    }
    i
}

fn is_single_line_if(stmt: &[char], ws: &[Word]) -> bool {
    let Some(then) = ws.iter().find(|w| w.text.eq_ignore_ascii_case("then")) else { return true };
    stmt.get(then.end..).is_some_and(|rest| rest.iter().any(|c| !c.is_whitespace()))
}

/// What a statement does to the block structure.
pub(super) fn classify(stmt: &[char]) -> Cls {
    let ws = words(stmt);
    let w: Vec<String> = ws.iter().map(|x| x.text.to_lowercase()).collect();
    let i = skip_modifiers(&w);
    let first = w.get(i).map(String::as_str).unwrap_or("");
    let next = w.get(i + 1).map(String::as_str).unwrap_or("");
    match first {
        "sub" | "function" => Cls::Open(Block::Proc),
        "property" if matches!(next, "get" | "let" | "set") => Cls::Open(Block::Proc),
        "type" if !next.is_empty() => Cls::Open(Block::Type),
        "enum" if !next.is_empty() => Cls::Open(Block::Enum),
        "if" if i == 0 => {
            if is_single_line_if(stmt, &ws) {
                Cls::Other
            } else {
                Cls::Open(Block::If)
            }
        }
        "elseif" if i == 0 => Cls::ElseIf,
        "else" if i == 0 => Cls::Else,
        "for" if i == 0 => Cls::Open(Block::For),
        "next" if i == 0 => {
            let vars = stmt.get(ws.first().map_or(0, |x| x.end)..).unwrap_or(&[]).iter().filter(|c| **c == ',').count();
            Cls::Close(Block::For, if next.is_empty() { 1 } else { vars + 1 })
        }
        "do" if i == 0 => Cls::Open(Block::Do),
        "loop" if i == 0 => Cls::Close(Block::Do, 1),
        "while" if i == 0 => Cls::Open(Block::While),
        "wend" if i == 0 => Cls::Close(Block::While, 1),
        "select" if i == 0 && next == "case" => Cls::Open(Block::Select),
        "case" if i == 0 => Cls::Case,
        "with" if i == 0 => Cls::Open(Block::With),
        "end" if i == 0 => match next {
            "sub" | "function" | "property" => Cls::Close(Block::Proc, 1),
            "if" => Cls::Close(Block::If, 1),
            "select" => Cls::Close(Block::Select, 1),
            "with" => Cls::Close(Block::With, 1),
            "type" => Cls::Close(Block::Type, 1),
            "enum" => Cls::Close(Block::Enum, 1),
            _ => Cls::Other,
        },
        _ => Cls::Other,
    }
}

// ---------------------------------------------------------------- declarations

fn first_lower(stmt: &[char]) -> String {
    words(stmt).first().map(|w| w.text.to_lowercase()).unwrap_or_default()
}

/// The type a name suffix stands for.
fn suffix_type(c: Option<char>) -> Option<&'static str> {
    match c? {
        '%' => Some("Integer"),
        '&' => Some("Long"),
        '!' => Some("Single"),
        '#' => Some("Double"),
        '@' => Some("Currency"),
        '$' => Some("String"),
        _ => None,
    }
}

/// Splits `text` at commas outside parentheses; each piece with its offset in `text`.
fn split_commas(text: &[char]) -> Vec<(usize, Vec<char>)> {
    let mut out = Vec::new();
    let (mut start, mut depth) = (0usize, 0i32);
    for (i, &c) in text.iter().enumerate() {
        match c {
            '(' => depth += 1,
            ')' => depth -= 1,
            ',' if depth == 0 => {
                out.push((start, text.get(start..i).unwrap_or(&[]).to_vec()));
                start = i + 1;
            }
            _ => {}
        }
    }
    out.push((start, text.get(start..).unwrap_or(&[]).to_vec()));
    out
}

/// The type after `As` in a declaration item (`As New Collection`, `As String * 10`).
fn declared_type(item: &[char]) -> Option<String> {
    let ws = words(item);
    let as_at = ws.iter().position(|w| w.text.eq_ignore_ascii_case("as"))?;
    let mut at = ws.get(as_at)?.end;
    let mut skip_new = false;
    if let Some(n) = ws.get(as_at + 1)
        && n.text.eq_ignore_ascii_case("new")
    {
        at = n.end;
        skip_new = true;
    }
    let rest: String = item.get(at..).unwrap_or(&[]).iter().collect();
    let name: String = rest.trim_start().chars().take_while(|c| is_word_char(*c) || *c == '.').collect();
    let _ = skip_new;
    if name.is_empty() { None } else { Some(name) }
}

/// The variables of a `Dim`/`Static`/`Public`/`Private`/`Const` statement.
fn declared_items(stmt: &[char], constant: bool) -> Vec<(String, Option<String>, usize)> {
    let ws = words(stmt);
    let i = ws
        .iter()
        .position(|w| matches!(w.text.to_lowercase().as_str(), "dim" | "const" | "static" | "public" | "private" | "global" | "friend"))
        .unwrap_or(0);
    // After the last leading keyword.
    let mut k = i;
    while ws
        .get(k + 1)
        .is_some_and(|w| matches!(w.text.to_lowercase().as_str(), "dim" | "const" | "static" | "withevents" | "public" | "private" | "global"))
    {
        k += 1;
    }
    let Some(kw) = ws.get(k) else { return Vec::new() };
    let rest = stmt.get(kw.end..).unwrap_or(&[]);
    let mut out = Vec::new();
    for (offset, item) in split_commas(rest) {
        let iw = words(&item);
        let Some(first) = iw.first() else { continue };
        let name_word = if first.text.eq_ignore_ascii_case("withevents") { iw.get(1) } else { Some(first) };
        let Some(name) = name_word else { continue };
        let ty = declared_type(&item).or_else(|| suffix_type(item.get(name.end).copied()).map(str::to_string));
        let ty = if constant && ty.is_none() { None } else { ty };
        out.push((name.text.clone(), ty, kw.end + offset + name.start));
    }
    out
}

fn parse_params(text: &[char]) -> Vec<ParamDecl> {
    split_commas(text)
        .into_iter()
        .filter_map(|(_, item)| {
            let ws = words(&item);
            let lower: Vec<String> = ws.iter().map(|w| w.text.to_lowercase()).collect();
            let name = ws.iter().zip(&lower).find(|(_, l)| !matches!(l.as_str(), "optional" | "byval" | "byref" | "paramarray")).map(|(w, _)| w)?;
            let default = item.iter().position(|c| *c == '=').map(|p| item.get(p + 1..).unwrap_or(&[]).iter().collect::<String>().trim().to_string());
            Some(ParamDecl {
                name: name.text.clone(),
                ty: declared_type(&item).or_else(|| suffix_type(item.get(name.end).copied()).map(str::to_string)),
                optional: lower.iter().any(|l| l == "optional"),
                by_val: lower.iter().any(|l| l == "byval"),
                param_array: lower.iter().any(|l| l == "paramarray"),
                default: default.filter(|d| !d.is_empty()),
            })
        })
        .collect()
}

struct Header {
    kind: ProcKind,
    public: bool,
    name: String,
    name_col: usize,
    params: Vec<ParamDecl>,
    returns: Option<String>,
}

fn parse_header(stmt: &[char]) -> Option<Header> {
    let ws = words(stmt);
    let w: Vec<String> = ws.iter().map(|x| x.text.to_lowercase()).collect();
    let i = skip_modifiers(&w);
    let public = !w.iter().take(i).any(|m| m == "private");
    let (kind, name_at) = match w.get(i).map(String::as_str)? {
        "sub" => (ProcKind::Sub, i + 1),
        "function" => (ProcKind::Function, i + 1),
        "property" => (
            match w.get(i + 1).map(String::as_str)? {
                "get" => ProcKind::PropertyGet,
                "let" => ProcKind::PropertyLet,
                _ => ProcKind::PropertySet,
            },
            i + 2,
        ),
        _ => return None,
    };
    let name = ws.get(name_at)?;
    let open = stmt.get(name.end..)?.iter().position(|c| *c == '(').map(|p| p + name.end);
    let (params, after) = match open {
        Some(o) if stmt.get(name.end..o)?.iter().all(|c| c.is_whitespace()) => {
            let mut depth = 0i32;
            let mut close = stmt.len();
            for (j, c) in stmt.iter().enumerate().skip(o) {
                match c {
                    '(' => depth += 1,
                    ')' => {
                        depth -= 1;
                        if depth == 0 {
                            close = j;
                            break;
                        }
                    }
                    _ => {}
                }
            }
            (parse_params(stmt.get(o + 1..close).unwrap_or(&[])), close.saturating_add(1))
        }
        _ => (Vec::new(), name.end),
    };
    let returns = declared_type(stmt.get(after..).unwrap_or(&[]));
    Some(Header { kind, public, name: name.text.clone(), name_col: name.start, params, returns })
}

// ---------------------------------------------------------------- the scan

struct Frame {
    block: Block,
    line: u32,
    start: usize,
    end: usize,
}

struct Scanner {
    scan: Scan,
    stack: Vec<Frame>,
    current_proc: Option<usize>,
    doc: Vec<String>,
    type_open: Option<usize>,
}

/// Reads `src`.
pub fn scan(src: &str) -> Scan {
    let lines = logical_lines(src);
    let mut s =
        Scanner { scan: Scan { lines: Vec::new(), ..Scan::default() }, stack: Vec::new(), current_proc: None, doc: Vec::new(), type_open: None };
    for line in &lines {
        s.line(line);
        s.scan.line_proc.push(s.current_proc);
    }
    s.finish();
    s.scan.lines = lines;
    s.scan
}

impl Scanner {
    fn error(&mut self, line: u32, start: usize, end: usize, template: &'static str, args: Vec<(&'static str, String)>) {
        self.scan.diagnostics.push(Diagnostic::new(line, start, end, Severity::Error, template, args));
    }

    fn line(&mut self, line: &Logical) {
        if line.blank {
            match &line.comment {
                Some(c) if line.first == line.last => self.doc.push(c.trim().to_string()),
                _ => self.doc.clear(),
            }
            return;
        }
        self.check_quotes(line);
        let code = &line.code;
        if code.iter().find(|c| !c.is_whitespace()) == Some(&'#') {
            self.doc.clear();
            return;
        }
        let stmts = statements(code);
        let Some((_, first)) = stmts.first() else { return };
        // A label starts in column one.
        let mut skip = 0;
        if line.indent == 0 && stmts.len() > 1
            || (line.indent == 0 && code.iter().rposition(|c| !c.is_whitespace()).is_some_and(|p| code.get(p) == Some(&':')))
        {
            let ws = words(first);
            if ws.len() == 1 && classify(first) == Cls::Other && !matches!(first_lower(first).as_str(), "else" | "loop" | "end") {
                let name = first_lower(first);
                self.scan.labels.push((ws.first().map_or(name, |w| w.text.clone()), line.first, self.current_proc));
                skip = 1;
            }
        }
        // `If x Then y: z` is one statement for the blocks.
        if classify(first) == Cls::Other && first_lower(first) == "if" {
            self.statement(line, 0, first);
            return;
        }
        for (offset, stmt) in stmts.iter().skip(skip) {
            self.statement(line, *offset, stmt);
        }
        self.doc.clear();
    }

    fn check_quotes(&mut self, line: &Logical) {
        // A string that is not closed: the masked code keeps its opening quote but loses the end.
        let quotes = line.code.iter().filter(|c| **c == '"').count();
        if quotes % 2 == 1 {
            let (l, col) = line.position(line.code.iter().rposition(|c| *c == '"').unwrap_or(0));
            self.error(l, col, line.first_line_end(), "Missing closing quote", vec![]);
        }
        let open = line.code.iter().filter(|c| **c == '(').count();
        let close = line.code.iter().filter(|c| **c == ')').count();
        if open != close {
            self.error(line.first, line.indent, line.first_line_end(), "Unbalanced parentheses", vec![]);
        }
    }

    fn statement(&mut self, line: &Logical, offset: usize, stmt: &[char]) {
        let (_, col) = line.position(offset);
        let lead = stmt.iter().take_while(|c| c.is_whitespace()).count();
        let (l, c0) = line.position(offset + lead);
        let _ = col;
        let (start, end) = (c0, line.first_line_end().max(c0 + 1));
        let cls = classify(stmt);
        let in_proc = self.stack.iter().any(|f| f.block == Block::Proc);
        match cls {
            Cls::Open(Block::Proc) => self.open_proc(line, l, c0, end, stmt),
            Cls::Open(b @ (Block::Type | Block::Enum)) => {
                if in_proc {
                    self.error(l, start, end, "Statements are not allowed outside a procedure", vec![]);
                }
                self.open_type(b, line, l, stmt);
                self.stack.push(Frame { block: b, line: l, start, end });
            }
            Cls::Open(b) => {
                if !in_proc {
                    self.error(l, start, end, "Statements are not allowed outside a procedure", vec![]);
                } else {
                    self.stack.push(Frame { block: b, line: l, start, end });
                }
            }
            Cls::Close(b, n) => {
                for _ in 0..n {
                    self.close(b, l, start, end);
                }
            }
            Cls::ElseIf | Cls::Else => {
                if self.stack.last().map(|f| f.block) != Some(Block::If) {
                    let word = if cls == Cls::Else { "Else" } else { "ElseIf" };
                    self.error(l, start, end, "“{word}” without a matching “{block}”", vec![("word", word.into()), ("block", "If".into())]);
                }
            }
            Cls::Case => {
                if self.stack.last().map(|f| f.block) != Some(Block::Select) {
                    self.error(
                        l,
                        start,
                        end,
                        "“{word}” without a matching “{block}”",
                        vec![("word", "Case".into()), ("block", "Select Case".into())],
                    );
                }
            }
            Cls::Other => self.other(line, l, start, end, stmt, in_proc),
        }
    }

    fn open_proc(&mut self, line: &Logical, l: u32, start: usize, end: usize, stmt: &[char]) {
        // A new procedure while something is open: the open blocks were never closed.
        let open: Vec<(Block, u32, usize, usize)> = self.stack.drain(..).map(|f| (f.block, f.line, f.start, f.end)).collect();
        for (b, bl, bs, be) in open {
            let kind = self.scan.procs.last().map(|p| p.kind);
            self.error(
                bl,
                bs,
                be,
                "Block not closed: “{open}” needs “{close}”",
                vec![("open", b.open_word().into()), ("close", b.close_word(kind).into())],
            );
        }
        if let Some(i) = self.current_proc.take()
            && let Some(p) = self.scan.procs.get_mut(i)
            && p.end_line == 0
        {
            p.end_line = 0;
        }
        let Some(h) = parse_header(stmt) else {
            self.error(l, start, end, "A procedure needs a name", vec![]);
            return;
        };
        let index = self.scan.procs.len();
        let doc = std::mem::take(&mut self.doc).join("\n");
        self.scan.procs.push(ProcDecl {
            name: h.name,
            kind: h.kind,
            public: h.public,
            line: l,
            end_line: 0,
            name_col: line.position(h.name_col).1.max(start),
            params: h.params.clone(),
            returns: h.returns,
            doc,
        });
        for p in &h.params {
            self.scan.vars.push(VarDecl {
                name: p.name.clone(),
                ty: p.ty.clone(),
                line: l,
                col: start,
                proc: Some(index),
                constant: false,
                public: false,
                param: true,
            });
        }
        self.current_proc = Some(index);
        self.stack.push(Frame { block: Block::Proc, line: l, start, end });
    }

    fn open_type(&mut self, b: Block, _line: &Logical, l: u32, stmt: &[char]) {
        let ws = words(stmt);
        let at = ws.iter().position(|w| matches!(w.text.to_lowercase().as_str(), "type" | "enum")).unwrap_or(0);
        if let Some(name) = ws.get(at + 1) {
            self.scan.types.push(TypeDecl { name: name.text.clone(), is_enum: b == Block::Enum, line: l, fields: Vec::new() });
            self.type_open = Some(self.scan.types.len() - 1);
        }
    }

    fn close(&mut self, b: Block, l: u32, start: usize, end: usize) {
        let found = self.stack.iter().rposition(|f| f.block == b);
        let Some(at) = found else {
            self.error(
                l,
                start,
                end,
                "“{close}” without a matching “{open}”",
                vec![("close", b.close_word(self.scan.procs.last().map(|p| p.kind)).into()), ("open", b.open_word().into())],
            );
            return;
        };
        // Blocks opened inside and never closed.
        let inner: Vec<Frame> = self.stack.drain(at + 1..).collect();
        for f in inner {
            let kind = self.scan.procs.last().map(|p| p.kind);
            self.error(
                f.line,
                f.start,
                f.end,
                "Block not closed: “{open}” needs “{close}”",
                vec![("open", f.block.open_word().into()), ("close", f.block.close_word(kind).into())],
            );
        }
        self.stack.pop();
        match b {
            Block::Proc => {
                if let Some(i) = self.current_proc.take()
                    && let Some(p) = self.scan.procs.get_mut(i)
                {
                    p.end_line = l;
                }
            }
            Block::Type | Block::Enum => self.type_open = None,
            _ => {}
        }
    }

    fn other(&mut self, line: &Logical, l: u32, start: usize, end: usize, stmt: &[char], in_proc: bool) {
        let first = first_lower(stmt);
        // Fields of a Type and members of an Enum.
        if let Some(t) = self.type_open
            && self.stack.last().is_some_and(|f| matches!(f.block, Block::Type | Block::Enum))
        {
            if let Some(w) = words(stmt).first() {
                let ty = declared_type(stmt);
                if let Some(decl) = self.scan.types.get_mut(t) {
                    decl.fields.push((w.text.clone(), ty));
                }
            }
            return;
        }
        let w: Vec<String> = words(stmt).iter().map(|x| x.text.to_lowercase()).collect();
        let i = skip_modifiers(&w);
        let head = w.get(i).map(String::as_str).unwrap_or("");
        let declares = matches!(head, "dim" | "const" | "withevents")
            || (i > 0 && !matches!(head, "declare" | "event" | "implements" | "sub" | "function" | "property"));
        if first == "option" && w.get(1).is_some_and(|x| x == "explicit") {
            self.scan.option_explicit = true;
            return;
        }
        if head == "declare" {
            if let Some(name) = declare_name(stmt) {
                self.scan.procs.push(ProcDecl {
                    name,
                    kind: if w.iter().any(|x| x == "function") { ProcKind::Function } else { ProcKind::Sub },
                    public: !w.iter().take(i).any(|m| m == "private"),
                    line: l,
                    end_line: l,
                    name_col: start,
                    params: Vec::new(),
                    returns: None,
                    doc: String::new(),
                });
            }
            return;
        }
        if declares || (first == "static" && in_proc) {
            let constant = w.iter().take(i + 2).any(|x| x == "const");
            let public = !in_proc && !w.iter().take(i).any(|m| m == "private" || m == "dim");
            for (name, ty, col) in declared_items(stmt, constant) {
                let (dl, dc) = line.position(col);
                self.scan.vars.push(VarDecl {
                    name,
                    ty,
                    line: dl,
                    col: dc,
                    proc: if in_proc { self.current_proc } else { None },
                    constant,
                    public,
                    param: false,
                });
            }
            return;
        }
        if first == "exit" {
            self.check_exit(l, start, end, &w);
            return;
        }
        if !in_proc
            && !matches!(
                first.as_str(),
                "attribute" | "event" | "implements" | "option" | "rem" | "defbool" | "defint" | "deflng" | "defstr" | "defdbl" | "defvar"
            )
        {
            self.error(l, start, end, "Statements are not allowed outside a procedure", vec![]);
        }
    }

    fn check_exit(&mut self, l: u32, start: usize, end: usize, w: &[String]) {
        let target = w.get(1).map(String::as_str).unwrap_or("");
        let kind = self.current_proc.and_then(|i| self.scan.procs.get(i)).map(|p| p.kind);
        let has = |b: Block| self.stack.iter().any(|f| f.block == b);
        let ok = match target {
            "sub" => kind == Some(ProcKind::Sub),
            "function" => kind == Some(ProcKind::Function),
            "property" => matches!(kind, Some(ProcKind::PropertyGet | ProcKind::PropertyLet | ProcKind::PropertySet)),
            "for" => has(Block::For),
            "do" => has(Block::Do),
            _ => true,
        };
        if !ok {
            let word = format!("Exit {}", w.get(1).map_or(String::new(), |t| capitalize(t)));
            self.error(l, start, end, "“{exit}” is not valid here", vec![("exit", word)]);
        }
    }

    fn finish(&mut self) {
        let open: Vec<Frame> = std::mem::take(&mut self.stack);
        for f in open {
            let kind = self.scan.procs.last().map(|p| p.kind);
            self.error(
                f.line,
                f.start,
                f.end,
                "Block not closed: “{open}” needs “{close}”",
                vec![("open", f.block.open_word().into()), ("close", f.block.close_word(kind).into())],
            );
        }
    }
}

fn capitalize(s: &str) -> String {
    let mut c = s.chars();
    c.next().map_or(String::new(), |f| f.to_uppercase().chain(c).collect())
}

/// The name in `Declare Function Name Lib "x" …`.
fn declare_name(stmt: &[char]) -> Option<String> {
    let ws = words(stmt);
    let at = ws.iter().position(|w| matches!(w.text.to_lowercase().as_str(), "function" | "sub"))?;
    ws.get(at + 1).map(|w| w.text.clone())
}

#[cfg(test)]
mod tests {
    use super::*;

    fn messages(src: &str) -> Vec<String> {
        scan(src).diagnostics.into_iter().map(|d| format!("{}:{}", d.line, d.message)).collect()
    }

    #[test]
    fn procedures_variables_and_parameters_are_found_without_parsing() {
        let s = scan(
            "Option Explicit\nPublic Counter As Long\nPrivate Const Limit = 5\n' Adds two numbers.\n' Really.\nPublic Function Add(ByVal a As Long, Optional b As Double = 2) As Double\n    Dim total As Double, name$\n    Add = a + b\nEnd Function\nSub Hello()\nEnd Sub\n",
        );
        assert!(s.option_explicit);
        assert_eq!(s.procs.len(), 2);
        let add = s.procs.first().unwrap();
        assert_eq!((add.name.as_str(), add.line, add.end_line), ("Add", 6, 9));
        assert_eq!(add.returns.as_deref(), Some("Double"));
        assert_eq!(add.doc, "Adds two numbers.\nReally.");
        assert_eq!(add.signature(), "Function Add(ByVal a As Long, [b As Double = 2]) As Double");
        let names: Vec<(&str, Option<&str>, bool)> = s.vars.iter().map(|v| (v.name.as_str(), v.ty.as_deref(), v.proc.is_none())).collect();
        assert!(names.contains(&("Counter", Some("Long"), true)));
        assert!(names.contains(&("Limit", None, true)));
        assert!(names.contains(&("total", Some("Double"), false)));
        assert!(names.contains(&("name", Some("String"), false)));
        assert!(names.contains(&("b", Some("Double"), false)));
        assert!(s.diagnostics.is_empty(), "{:?}", s.diagnostics);
    }

    #[test]
    fn blocks_that_do_not_match_are_reported_where_they_open() {
        let m = messages("Sub A()\n    If x Then\n        y = 1\nEnd Sub\n");
        assert_eq!(m, ["2:Block not closed: “If” needs “End If”"]);
        let m = messages("Sub A()\n    For i = 1 To 3\n    Next\n    Next\nEnd Sub\n");
        assert_eq!(m, ["4:“Next” without a matching “For”"]);
        let m = messages("Sub A()\nEnd Sub\nEnd Sub\n");
        assert_eq!(m, ["3:“End Sub” without a matching “Sub”"]);
        let m = messages("Sub A()\n    Do While x\n        y = 1\n");
        assert_eq!(m.len(), 2, "{m:?}");
        assert!(m.iter().any(|x| x.contains("“Do” needs “Loop”")));
        assert!(m.iter().any(|x| x.contains("“Sub” needs “End Sub”")));
    }

    #[test]
    fn a_missing_end_does_not_swallow_the_next_procedure() {
        let s = scan("Sub A()\n    x = 1\nSub B()\nEnd Sub\n");
        assert_eq!(s.procs.len(), 2);
        assert_eq!(s.procs.first().unwrap().end_line, 0);
        assert_eq!(s.procs.get(1).unwrap().end_line, 4);
        assert_eq!(s.diagnostics.len(), 1);
    }

    #[test]
    fn single_line_if_and_colons_do_not_open_blocks() {
        assert!(
            messages("Sub A()\n    If x Then y = 1 Else y = 2\n    If a Then b = 1: c = 2\n    For i = 1 To 2: x = i: Next i\nEnd Sub\n").is_empty()
        );
    }

    #[test]
    fn select_case_and_with_and_else_are_checked() {
        assert!(messages("Sub A()\n    Select Case x\n    Case 1\n        y = 1\n    Case Else\n        y = 2\n    End Select\n    With r\n        .A = 1\n    End With\n    If a Then\n    ElseIf b Then\n    Else\n    End If\nEnd Sub\n").is_empty());
        let m = messages("Sub A()\n    Else\n    Case 1\nEnd Sub\n");
        assert_eq!(m.len(), 2, "{m:?}");
    }

    #[test]
    fn statements_outside_procedures_and_bad_exits_are_reported() {
        let m = messages("x = 1\nSub A()\n    Exit Function\nEnd Sub\n");
        assert_eq!(m, ["1:Statements are not allowed outside a procedure", "3:“Exit Function” is not valid here"]);
    }

    #[test]
    fn labels_types_enums_and_declares_are_understood() {
        let s = scan(
            "Private Type Person\n    Name As String\n    Age As Integer\nEnd Type\nPublic Enum Color\n    Red = 1\n    Green\nEnd Enum\nPrivate Declare Function GetTick Lib \"kernel32\" () As Long\nSub A()\n    On Error GoTo Done\n    x = 1\nDone:\n    Exit Sub\nEnd Sub\n",
        );
        assert!(s.diagnostics.is_empty(), "{:?}", s.diagnostics);
        let person = s.types.first().unwrap();
        assert_eq!(person.fields, [("Name".to_string(), Some("String".to_string())), ("Age".to_string(), Some("Integer".to_string()))]);
        assert!(s.types.get(1).unwrap().is_enum);
        assert_eq!(s.types.get(1).unwrap().fields.len(), 2);
        assert!(s.procs.iter().any(|p| p.name == "GetTick"));
        assert_eq!(s.labels.first().map(|l| l.0.as_str()), Some("Done"));
    }

    #[test]
    fn broken_quotes_and_parentheses_are_reported() {
        assert_eq!(messages("Sub A()\n    x = \"abc\nEnd Sub\n"), ["2:Missing closing quote"]);
        assert_eq!(messages("Sub A()\n    x = f(1\nEnd Sub\n"), ["2:Unbalanced parentheses"]);
    }

    #[test]
    fn hostile_text_does_not_panic() {
        for src in ["", "\u{0}", "Sub", "End", "Property", "Sub (", "Dim", "Next , , ,", "If Then", "\"", "_\n_", "Function f(", "Type"] {
            let _ = scan(src);
        }
        let big = "Sub A()\n".repeat(2000);
        assert!(scan(&big).procs.len() == 2000);
    }
}
