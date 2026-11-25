//! Tokens to a [`Module`]: a recursive-descent parser for the VBA statements and expressions
//! macros use.

use std::rc::Rc;

#[allow(clippy::wildcard_imports)]
use crate::ast::*;
use crate::lexer::{SyntaxError, Tok, Token, lex};
use crate::value::Value;

type PResult<T> = Result<T, SyntaxError>;

/// Deepest nesting of expressions and blocks the parser follows.
const MAX_DEPTH: u32 = 150;

pub fn parse_module(name: &str, src: &str) -> PResult<Module> {
    let tokens = lex(src)?;
    let mut p = Parser::new(tokens);
    let mut module = Module { name: name.to_string(), ..Module::default() };
    p.module(&mut module)?;
    Ok(module)
}

/// Statements typed on their own (the Immediate window), numbered from line 1.
pub(crate) fn parse_snippet(src: &str) -> PResult<Vec<Stmt>> {
    let mut p = Parser::new(lex(src)?);
    let out = p.block()?;
    if p.peek().is_some() {
        return p.err("unexpected text");
    }
    Ok(out)
}

/// One expression typed on its own (a watch).
pub(crate) fn parse_expression(src: &str) -> PResult<Expr> {
    let mut p = Parser::new(lex(src)?);
    let e = p.expr()?;
    p.skip_eols();
    if p.peek().is_some() {
        return p.err("unexpected text after the expression");
    }
    Ok(e)
}

struct Parser {
    toks: Vec<Token>,
    /// Each identifier as it was spelled, for the names a person sees (`toks` holds them lower case).
    spelled: Vec<Option<Rc<str>>>,
    pos: usize,
    depth: u32,
}

/// Keywords that end a block; the construct that opened it checks which one it got.
fn is_block_end_word(w: &str) -> bool {
    matches!(w, "else" | "elseif" | "next" | "loop" | "wend" | "case")
}

impl Parser {
    fn new(mut toks: Vec<Token>) -> Parser {
        let mut spelled = Vec::with_capacity(toks.len());
        for t in &mut toks {
            if let Tok::Ident(n, _) = &mut t.tok {
                spelled.push(Some(n.clone()));
                *n = Rc::from(n.to_lowercase().as_str());
            } else {
                spelled.push(None);
            }
        }
        Parser { toks, spelled, pos: 0, depth: 0 }
    }

    // ---- token access ----

    fn peek(&self) -> Option<&Tok> {
        self.toks.get(self.pos).map(|t| &t.tok)
    }

    fn peek_at(&self, n: usize) -> Option<&Tok> {
        self.toks.get(self.pos + n).map(|t| &t.tok)
    }

    fn line(&self) -> u32 {
        self.toks.get(self.pos).or_else(|| self.toks.last()).map_or(0, |t| t.line)
    }

    fn err<T>(&self, message: impl Into<String>) -> PResult<T> {
        Err(SyntaxError { line: self.line(), message: message.into() })
    }

    fn advance(&mut self) {
        self.pos += 1;
    }

    fn at_eol(&self) -> bool {
        matches!(self.peek(), None | Some(Tok::Eol))
    }

    /// The current token as a keyword (lower case), when it is an identifier.
    fn word(&self) -> Option<&str> {
        match self.peek() {
            Some(Tok::Ident(n, _)) => Some(n),
            _ => None,
        }
    }

    fn word_at(&self, n: usize) -> Option<&str> {
        match self.peek_at(n) {
            Some(Tok::Ident(w, _)) => Some(w),
            _ => None,
        }
    }

    fn is_word(&self, w: &str) -> bool {
        self.word() == Some(w)
    }

    fn accept_word(&mut self, w: &str) -> bool {
        if self.is_word(w) {
            self.advance();
            true
        } else {
            false
        }
    }

    fn expect_word(&mut self, w: &str) -> PResult<()> {
        if self.accept_word(w) { Ok(()) } else { self.err(format!("expected `{w}`")) }
    }

    fn is_p(&self, p: &str) -> bool {
        matches!(self.peek(), Some(Tok::P(q)) if *q == p)
    }

    fn accept_p(&mut self, p: &str) -> bool {
        if self.is_p(p) {
            self.advance();
            true
        } else {
            false
        }
    }

    fn expect_p(&mut self, p: &str) -> PResult<()> {
        if self.accept_p(p) { Ok(()) } else { self.err(format!("expected `{p}`")) }
    }

    fn ident(&mut self) -> PResult<Name> {
        match self.peek() {
            Some(Tok::Ident(n, _)) => {
                let n = n.clone();
                self.advance();
                Ok(n)
            }
            _ => self.err("expected a name"),
        }
    }

    fn skip_eols(&mut self) {
        while matches!(self.peek(), Some(Tok::Eol)) {
            self.advance();
        }
    }

    fn skip_line(&mut self) {
        while !self.at_eol() {
            self.advance();
        }
    }

    fn end_of_statement(&mut self) -> PResult<()> {
        if self.at_eol() {
            self.skip_eols();
            return Ok(());
        }
        if self.accept_p(":") {
            return Ok(());
        }
        if self.is_word("else") {
            return Ok(());
        }
        self.err("unexpected text at the end of the statement")
    }

    fn enter(&mut self) -> PResult<()> {
        self.depth += 1;
        if self.depth > MAX_DEPTH { self.err("the code is nested too deeply") } else { Ok(()) }
    }

    fn leave(&mut self) {
        self.depth = self.depth.saturating_sub(1);
    }

    // ---- module level ----

    fn module(&mut self, m: &mut Module) -> PResult<()> {
        loop {
            self.skip_eols();
            if self.peek().is_none() {
                return Ok(());
            }
            self.declaration(m)?;
        }
    }

    fn declaration(&mut self, m: &mut Module) -> PResult<()> {
        let mut public = true;
        let mut is_static = false;
        loop {
            match self.word() {
                Some("public" | "global") => public = true,
                Some("private" | "friend") => public = false,
                Some("static") => is_static = true,
                _ => break,
            }
            self.advance();
        }
        let _ = is_static;
        match self.word() {
            Some(
                "option" | "attribute" | "implements" | "event" | "declare" | "deftype" | "defint" | "deflng" | "defdbl" | "defstr" | "defbool"
                | "defvar",
            ) => {
                self.skip_line();
                Ok(())
            }
            Some("sub" | "function" | "property") => {
                let p = self.procedure(public)?;
                m.procs.push(p);
                Ok(())
            }
            Some("dim") | Some("const") => self.module_decl(m),
            Some("type") => self.skip_until_end("type"),
            Some("enum") => self.enumeration(m),
            Some(_) if self.word_at(1).is_some() || self.peek_at(1) == Some(&Tok::P("(")) || self.peek_at(1) == Some(&Tok::P("=")) => {
                // `Public x As Long` after the modifier was taken.
                self.module_decl(m)
            }
            _ => self.err("a statement outside of any procedure"),
        }
    }

    fn skip_until_end(&mut self, what: &str) -> PResult<()> {
        loop {
            if self.peek().is_none() {
                return self.err(format!("`End {what}` is missing"));
            }
            if self.is_word("end") && self.word_at(1) == Some(what) {
                self.skip_line();
                return Ok(());
            }
            self.advance();
        }
    }

    fn enumeration(&mut self, m: &mut Module) -> PResult<()> {
        self.advance();
        self.skip_line();
        let mut next = 0i32;
        loop {
            self.skip_eols();
            if self.peek().is_none() {
                return self.err("`End Enum` is missing");
            }
            if self.is_word("end") {
                self.skip_line();
                return Ok(());
            }
            let name = self.ident()?;
            if self.accept_p("=") {
                let e = self.expr()?;
                if let Expr::Lit(v) = &e
                    && let Ok(n) = v.to_i32()
                {
                    next = n;
                }
                m.decls.push(Decl::Const(name, e));
            } else {
                m.decls.push(Decl::Const(name, Expr::Lit(Value::Long(next))));
            }
            next = next.saturating_add(1);
            self.skip_line();
        }
    }

    fn module_decl(&mut self, m: &mut Module) -> PResult<()> {
        if self.accept_word("const") {
            for (n, e) in self.const_items()? {
                m.decls.push(Decl::Const(n, e));
            }
        } else {
            self.accept_word("dim");
            for item in self.dim_items()? {
                m.decls.push(Decl::Var(item));
            }
        }
        self.end_of_statement()
    }

    fn procedure(&mut self, public: bool) -> PResult<Proc> {
        let line = self.line();
        let kind = match self.word() {
            Some("sub") => ProcKind::Sub,
            Some("function") => ProcKind::Function,
            _ => {
                self.advance();
                match self.word() {
                    Some("get") => ProcKind::PropertyGet,
                    Some("let") => ProcKind::PropertyLet,
                    Some("set") => ProcKind::PropertySet,
                    _ => return self.err("expected Get, Let or Set"),
                }
            }
        };
        self.advance();
        let display = match (self.peek(), self.spelled.get(self.pos)) {
            (Some(Tok::Ident(n, _)), Some(spelled)) => spelled.as_ref().map_or_else(|| n.to_string(), |s| s.to_string()),
            _ => return self.err("expected the procedure's name"),
        };
        let name = self.ident()?;
        let params = if self.is_p("(") { self.params()? } else { Vec::new() };
        let returns = if self.accept_word("as") { self.type_name()? } else { VType::Variant };
        self.skip_line();
        self.skip_eols();
        let body = self.block()?;
        if !self.is_word("end") {
            return self.err(format!("`End {}` is missing", if kind == ProcKind::Sub { "Sub" } else { "Function" }));
        }
        self.skip_line();
        Ok(Proc { name, display, kind, public, params, returns, body, line })
    }

    fn params(&mut self) -> PResult<Vec<Param>> {
        self.expect_p("(")?;
        let mut out = Vec::new();
        while !self.is_p(")") {
            out.push(self.param()?);
            if !self.accept_p(",") {
                break;
            }
        }
        self.expect_p(")")?;
        Ok(out)
    }

    fn param(&mut self) -> PResult<Param> {
        let (mut optional, mut by_val, mut param_array) = (false, false, false);
        loop {
            match self.word() {
                Some("optional") => optional = true,
                Some("byval") => by_val = true,
                Some("byref") => by_val = false,
                Some("paramarray") => param_array = true,
                _ => break,
            }
            self.advance();
        }
        let name = self.ident()?;
        if self.accept_p("(") {
            self.expect_p(")")?;
        }
        let ty = if self.accept_word("as") { self.type_name()? } else { VType::Variant };
        let default = if self.accept_p("=") { Some(self.expr()?) } else { None };
        Ok(Param { name, by_val, optional, param_array, ty, default })
    }

    /// `Long`, `Excel.Range`, `String * 10`.
    fn type_name(&mut self) -> PResult<VType> {
        let mut name = self.ident()?;
        while self.is_p(".") {
            self.advance();
            name = self.ident()?;
        }
        if self.is_p("*") {
            self.advance();
            self.advance();
        }
        Ok(VType::from_name(&name))
    }

    // ---- declarations ----

    fn const_items(&mut self) -> PResult<Vec<(Name, Expr)>> {
        let mut out = Vec::new();
        loop {
            let name = self.ident()?;
            if self.accept_word("as") {
                self.type_name()?;
            }
            self.expect_p("=")?;
            out.push((name, self.expr()?));
            if !self.accept_p(",") {
                return Ok(out);
            }
        }
    }

    fn dim_items(&mut self) -> PResult<Vec<DimItem>> {
        self.accept_word("withevents");
        let mut out = Vec::new();
        loop {
            out.push(self.dim_item()?);
            if !self.accept_p(",") {
                return Ok(out);
            }
        }
    }

    fn dim_item(&mut self) -> PResult<DimItem> {
        let name = self.ident()?;
        let (mut bounds, mut dynamic) = (None, false);
        if self.accept_p("(") {
            if self.accept_p(")") {
                dynamic = true;
            } else {
                let mut dims = Vec::new();
                loop {
                    let first = self.expr()?;
                    if self.accept_word("to") {
                        dims.push((Some(first), self.expr()?));
                    } else {
                        dims.push((None, first));
                    }
                    if !self.accept_p(",") {
                        break;
                    }
                }
                self.expect_p(")")?;
                bounds = Some(dims);
            }
        }
        let (mut ty, mut is_new) = (VType::Variant, false);
        if self.accept_word("as") {
            is_new = self.accept_word("new");
            ty = self.type_name()?;
        }
        Ok(DimItem { name, ty, bounds, dynamic, is_new })
    }

    // ---- statements ----

    /// Statements up to a block-ending keyword (left unconsumed).
    fn block(&mut self) -> PResult<Vec<Stmt>> {
        self.enter()?;
        let mut out = Vec::new();
        loop {
            self.skip_eols();
            while self.accept_p(":") {}
            if self.peek().is_none() || self.at_block_end() {
                break;
            }
            let line = self.line();
            let kind = self.statement()?;
            out.push(Stmt { line, kind });
            self.after_statement(false)?;
        }
        self.leave();
        Ok(out)
    }

    fn at_block_end(&self) -> bool {
        match self.word() {
            Some(w) if is_block_end_word(w) => true,
            Some("end") => matches!(self.word_at(1), Some("if" | "sub" | "function" | "select" | "with" | "property" | "type" | "enum")),
            _ => false,
        }
    }

    /// Statements of a single-line `If` branch: up to the end of the line or an `Else`.
    fn inline_statements(&mut self) -> PResult<Vec<Stmt>> {
        let mut out = Vec::new();
        loop {
            if self.at_eol() || self.is_word("else") {
                return Ok(out);
            }
            let line = self.line();
            let kind = self.statement()?;
            out.push(Stmt { line, kind });
            self.after_statement(true)?;
        }
    }

    /// What may follow a statement: the end of the line, a `:` (and then more statements), or in
    /// a single-line `If` an `Else`.
    fn after_statement(&mut self, inline: bool) -> PResult<()> {
        if self.accept_p(":") || self.at_eol() || (inline && self.is_word("else")) {
            return Ok(());
        }
        self.err("unexpected text at the end of the statement")
    }

    fn statement(&mut self) -> PResult<StmtKind> {
        if self.at_label() {
            let name = self.ident()?;
            self.advance();
            return Ok(StmtKind::Label(name));
        }
        let kw = self.word().map(str::to_string);
        // `line = 3` assigns to a variable that happens to be named like a keyword.
        let keyword_use = !matches!(self.peek_at(1), Some(Tok::P("=")));
        if let Some(k) = kw
            && keyword_use
            && let Some(stmt) = self.keyword_statement(&k)?
        {
            return Ok(stmt);
        }
        self.expression_statement()
    }

    fn at_label(&self) -> bool {
        matches!(self.peek(), Some(Tok::Ident(..)))
            && matches!(self.peek_at(1), Some(Tok::P(":")))
            && !matches!(self.word(), Some("case" | "else" | "default"))
            && !matches!(self.peek_at(2), Some(Tok::P("=")))
    }

    fn keyword_statement(&mut self, k: &str) -> PResult<Option<StmtKind>> {
        Ok(Some(match k {
            "dim" | "static" => {
                self.advance();
                StmtKind::Dim { items: self.dim_items()?, is_static: k == "static" }
            }
            "const" => {
                self.advance();
                StmtKind::Const(self.const_items()?)
            }
            "redim" => {
                self.advance();
                let preserve = self.accept_word("preserve");
                StmtKind::ReDim { preserve, items: self.dim_items()? }
            }
            "set" => {
                self.advance();
                let target = self.postfix()?;
                self.expect_p("=")?;
                StmtKind::SetAssign(target, self.expr()?)
            }
            "let" => {
                self.advance();
                let target = self.postfix()?;
                self.expect_p("=")?;
                StmtKind::Assign(target, self.expr()?)
            }
            "call" => {
                self.advance();
                StmtKind::Call(self.postfix()?)
            }
            "if" => self.if_statement()?,
            "for" => self.for_statement()?,
            "do" => self.do_statement()?,
            "while" => {
                self.advance();
                let cond = self.expr()?;
                let body = self.block()?;
                self.expect_word("wend")?;
                StmtKind::While(cond, body)
            }
            "select" => self.select_statement()?,
            "with" => {
                self.advance();
                let target = self.expr()?;
                let body = self.block()?;
                self.end_block("with")?;
                StmtKind::With(target, body)
            }
            "exit" => self.exit_statement()?,
            "on" => self.on_statement()?,
            "goto" => {
                self.advance();
                StmtKind::Goto(self.ident()?)
            }
            "resume" => {
                self.advance();
                let next = self.accept_word("next");
                if !next && !self.at_eol() {
                    self.advance();
                }
                StmtKind::Resume(next)
            }
            "open" => self.open_statement()?,
            "close" => {
                self.advance();
                let mut v = Vec::new();
                while !self.at_eol() && !self.is_p(":") {
                    self.accept_p("#");
                    v.push(self.expr()?);
                    if !self.accept_p(",") {
                        break;
                    }
                }
                StmtKind::Close(v)
            }
            "print" | "write" if matches!(self.peek_at(1), Some(Tok::P("#"))) => self.print_statement(k == "write")?,
            "line" if self.word_at(1) == Some("input") => {
                self.advance();
                self.advance();
                self.accept_p("#");
                let number = self.expr()?;
                self.expect_p(",")?;
                StmtKind::LineInput { number, target: self.postfix()? }
            }
            "erase" => {
                self.advance();
                let mut v = vec![self.postfix()?];
                while self.accept_p(",") {
                    v.push(self.postfix()?);
                }
                StmtKind::Erase(v)
            }
            "end" if matches!(self.peek_at(1), None | Some(Tok::Eol) | Some(Tok::P(":"))) => {
                self.advance();
                StmtKind::End
            }
            "stop" => {
                self.advance();
                StmtKind::Stop
            }
            "randomize" | "doevents" if self.at_end_after_word() => {
                self.advance();
                StmtKind::Nop
            }
            _ => return Ok(None),
        }))
    }

    fn at_end_after_word(&self) -> bool {
        matches!(self.peek_at(1), None | Some(Tok::Eol) | Some(Tok::P(":")))
    }

    fn end_block(&mut self, what: &str) -> PResult<()> {
        if self.is_word("end") && self.word_at(1) == Some(what) {
            self.advance();
            self.advance();
            Ok(())
        } else {
            self.err(format!("`End {what}` is missing"))
        }
    }

    fn if_statement(&mut self) -> PResult<StmtKind> {
        self.advance();
        let cond = self.expr()?;
        self.expect_word("then")?;
        if !self.at_eol() {
            // Single-line form.
            let then = self.inline_statements()?;
            let otherwise = if self.accept_word("else") { Some(self.inline_statements()?) } else { None };
            return Ok(StmtKind::If { branches: vec![(cond, then)], otherwise });
        }
        let mut branches = vec![(cond, self.block()?)];
        let mut otherwise = None;
        loop {
            if self.accept_word("elseif") {
                let c = self.expr()?;
                self.expect_word("then")?;
                branches.push((c, self.block()?));
            } else if self.accept_word("else") {
                if self.word() == Some("if") {
                    // `Else If` on one line behaves like `ElseIf`.
                    self.advance();
                    let c = self.expr()?;
                    self.expect_word("then")?;
                    branches.push((c, self.block()?));
                    continue;
                }
                otherwise = Some(self.block()?);
            } else {
                break;
            }
        }
        self.end_block("if")?;
        Ok(StmtKind::If { branches, otherwise })
    }

    fn for_statement(&mut self) -> PResult<StmtKind> {
        self.advance();
        if self.accept_word("each") {
            let var = self.postfix()?;
            self.expect_word("in")?;
            let collection = self.expr()?;
            let body = self.block()?;
            self.finish_next()?;
            return Ok(StmtKind::ForEach { var, collection, body });
        }
        let var = self.postfix()?;
        self.expect_p("=")?;
        let from = self.expr()?;
        self.expect_word("to")?;
        let to = self.expr()?;
        let step = if self.accept_word("step") { Some(self.expr()?) } else { None };
        let body = self.block()?;
        self.finish_next()?;
        Ok(StmtKind::For { var, from, to, step, body })
    }

    /// `Next [i[, j]]`: the counters are optional and ignored.
    fn finish_next(&mut self) -> PResult<()> {
        self.expect_word("next")?;
        while !self.at_eol() && !self.is_p(":") {
            self.advance();
        }
        Ok(())
    }

    fn do_statement(&mut self) -> PResult<StmtKind> {
        self.advance();
        let mut cond = None;
        let mut until = false;
        let mut post = false;
        if self.is_word("while") || self.is_word("until") {
            until = self.is_word("until");
            self.advance();
            cond = Some(self.expr()?);
        }
        let body = self.block()?;
        self.expect_word("loop")?;
        if cond.is_none() && (self.is_word("while") || self.is_word("until")) {
            until = self.is_word("until");
            self.advance();
            cond = Some(self.expr()?);
            post = true;
        }
        Ok(StmtKind::Do { cond, until, post, body })
    }

    fn select_statement(&mut self) -> PResult<StmtKind> {
        self.advance();
        self.expect_word("case")?;
        let subject = self.expr()?;
        self.skip_line();
        self.skip_eols();
        let mut cases = Vec::new();
        let mut otherwise = None;
        while self.is_word("case") {
            self.advance();
            if self.accept_word("else") {
                self.accept_p(":");
                otherwise = Some(self.block()?);
                continue;
            }
            let tests = self.case_tests()?;
            self.accept_p(":");
            cases.push((tests, self.block()?));
        }
        self.end_block("select")?;
        Ok(StmtKind::Select { subject, cases, otherwise })
    }

    fn case_tests(&mut self) -> PResult<Vec<CaseTest>> {
        let mut out = Vec::new();
        loop {
            if self.accept_word("is") {
                let op = match self.peek() {
                    Some(Tok::P("=")) => Bin::Eq,
                    Some(Tok::P("<>")) => Bin::Ne,
                    Some(Tok::P("<")) => Bin::Lt,
                    Some(Tok::P(">")) => Bin::Gt,
                    Some(Tok::P("<=")) => Bin::Le,
                    Some(Tok::P(">=")) => Bin::Ge,
                    _ => return self.err("expected a comparison after `Is`"),
                };
                self.advance();
                out.push(CaseTest::Is(op, self.expr()?));
            } else {
                let first = self.expr()?;
                if self.accept_word("to") {
                    out.push(CaseTest::Range(first, self.expr()?));
                } else {
                    out.push(CaseTest::Value(first));
                }
            }
            if !self.accept_p(",") {
                return Ok(out);
            }
        }
    }

    fn exit_statement(&mut self) -> PResult<StmtKind> {
        self.advance();
        let kind = match self.word() {
            Some("sub") => ExitKind::Sub,
            Some("function") => ExitKind::Function,
            Some("property") => ExitKind::Property,
            Some("for") => ExitKind::For,
            Some("do") => ExitKind::Do,
            _ => return self.err("expected Sub, Function, For or Do after Exit"),
        };
        self.advance();
        Ok(StmtKind::Exit(kind))
    }

    fn on_statement(&mut self) -> PResult<StmtKind> {
        self.advance();
        self.expect_word("error")?;
        if self.accept_word("resume") {
            self.expect_word("next")?;
            return Ok(StmtKind::OnError(OnError::ResumeNext));
        }
        self.expect_word("goto")?;
        match self.peek() {
            Some(Tok::Num(v)) if v.as_number() == Some(0.0) => {
                self.advance();
                Ok(StmtKind::OnError(OnError::Goto0))
            }
            Some(Tok::P("-")) => {
                // On Error GoTo -1
                self.advance();
                self.advance();
                Ok(StmtKind::OnError(OnError::Goto0))
            }
            _ => Ok(StmtKind::OnError(OnError::Goto(self.ident()?))),
        }
    }

    fn open_statement(&mut self) -> PResult<StmtKind> {
        self.advance();
        let path = self.expr()?;
        self.expect_word("for")?;
        let mode = match self.word() {
            Some("input") => FileMode::Input,
            Some("output") => FileMode::Output,
            Some("append") => FileMode::Append,
            Some("binary") => FileMode::Binary,
            Some("random") => FileMode::Random,
            _ => return self.err("expected Input, Output or Append"),
        };
        self.advance();
        // Access / Lock clauses.
        while !self.is_word("as") && !self.at_eol() {
            self.advance();
        }
        self.expect_word("as")?;
        self.accept_p("#");
        let number = self.expr()?;
        Ok(StmtKind::Open { path, mode, number })
    }

    fn print_statement(&mut self, quoted: bool) -> PResult<StmtKind> {
        self.advance();
        self.expect_p("#")?;
        let number = self.expr()?;
        let mut items = Vec::new();
        let mut newline = true;
        if self.accept_p(",") {
            while !self.at_eol() && !self.is_p(":") && !self.is_word("else") {
                newline = true;
                if self.is_p(";") || self.is_p(",") {
                    // A separator with nothing before it: Print's tab zones are not modelled.
                    self.advance();
                    newline = false;
                    continue;
                }
                items.push(self.expr()?);
                if self.accept_p(";") || self.accept_p(",") {
                    newline = false;
                }
            }
        }
        Ok(StmtKind::Print { number, items, newline, quoted })
    }

    /// An assignment or a call written as a statement.
    fn expression_statement(&mut self) -> PResult<StmtKind> {
        let target = self.postfix()?;
        if self.accept_p("=") {
            return Ok(StmtKind::Assign(target, self.expr()?));
        }
        let callable = matches!(target, Expr::Name(_) | Expr::Member(..) | Expr::WithMember(_) | Expr::Call(..));
        if !callable {
            return self.err("a statement was expected");
        }
        if self.at_eol() || self.is_p(":") || self.is_word("else") {
            return Ok(StmtKind::Call(target));
        }
        // `Name arg, arg`: a call without parentheses.
        let args = self.command_args()?;
        Ok(StmtKind::Call(attach_args(target, args)))
    }

    fn command_args(&mut self) -> PResult<Vec<Arg>> {
        let mut out = Vec::new();
        loop {
            if self.is_p(",") {
                out.push(Arg { name: None, value: None });
                self.advance();
                continue;
            }
            out.push(self.arg()?);
            if !self.accept_p(",") {
                return Ok(out);
            }
        }
    }

    fn arg(&mut self) -> PResult<Arg> {
        if matches!(self.peek(), Some(Tok::Ident(..))) && matches!(self.peek_at(1), Some(Tok::P(":="))) {
            let name = self.ident()?;
            self.advance();
            return Ok(Arg { name: Some(name), value: Some(self.expr()?) });
        }
        Ok(Arg { name: None, value: Some(self.expr()?) })
    }

    fn paren_args(&mut self) -> PResult<Vec<Arg>> {
        self.expect_p("(")?;
        let mut out = Vec::new();
        if self.accept_p(")") {
            return Ok(out);
        }
        loop {
            if self.is_p(",") || self.is_p(")") {
                out.push(Arg { name: None, value: None });
            } else {
                out.push(self.arg()?);
            }
            if self.accept_p(",") {
                if self.is_p(")") {
                    out.push(Arg { name: None, value: None });
                }
                continue;
            }
            break;
        }
        self.expect_p(")")?;
        Ok(out)
    }

    // ---- expressions ----

    pub fn expr(&mut self) -> PResult<Expr> {
        self.enter()?;
        let e = self.binary(0);
        self.leave();
        e
    }

    /// The operator at the cursor and its precedence level (higher binds tighter).
    fn operator(&self) -> Option<(Bin, u8)> {
        Some(match self.peek()? {
            Tok::P(p) => match *p {
                "^" => (Bin::Pow, 12),
                "*" => (Bin::Mul, 10),
                "/" => (Bin::Div, 10),
                "\\" => (Bin::IntDiv, 9),
                "+" => (Bin::Add, 7),
                "-" => (Bin::Sub, 7),
                "&" => (Bin::Concat, 6),
                "=" => (Bin::Eq, 5),
                "<>" => (Bin::Ne, 5),
                "<" => (Bin::Lt, 5),
                ">" => (Bin::Gt, 5),
                "<=" => (Bin::Le, 5),
                ">=" => (Bin::Ge, 5),
                _ => return None,
            },
            Tok::Ident(n, _) => match &**n {
                "mod" => (Bin::Mod, 8),
                "like" => (Bin::Like, 5),
                "is" => (Bin::Is, 5),
                "and" => (Bin::And, 3),
                "or" => (Bin::Or, 2),
                "xor" => (Bin::Xor, 2),
                "eqv" => (Bin::Eqv, 1),
                "imp" => (Bin::Imp, 0),
                _ => return None,
            },
            _ => return None,
        })
    }

    /// Precedence climbing; `min` is the lowest level this call accepts.
    fn binary(&mut self, min: u8) -> PResult<Expr> {
        let mut lhs = self.unary()?;
        while let Some((op, level)) = self.operator() {
            if level < min {
                break;
            }
            self.advance();
            // `^` and the others are left-associative in VBA.
            let rhs = self.binary(level + 1)?;
            lhs = Expr::Binary(op, Box::new(lhs), Box::new(rhs));
        }
        Ok(lhs)
    }

    fn unary(&mut self) -> PResult<Expr> {
        if self.is_word("not") {
            self.advance();
            // `Not` binds looser than comparison but tighter than `And`.
            let e = self.binary(4)?;
            return Ok(Expr::Unary(Un::Not, Box::new(e)));
        }
        if self.is_p("-") || self.is_p("+") {
            let neg = self.is_p("-");
            self.advance();
            // Unary minus binds tighter than `*` but looser than `^`.
            let e = self.binary(11)?;
            return Ok(Expr::Unary(if neg { Un::Neg } else { Un::Plus }, Box::new(e)));
        }
        self.postfix()
    }

    /// A primary expression followed by member accesses and argument lists.
    fn postfix(&mut self) -> PResult<Expr> {
        let mut e = self.primary()?;
        loop {
            if self.is_p(".") && !matches!(self.peek_at(1), Some(Tok::Num(_))) {
                self.advance();
                let name = self.member_name()?;
                e = Expr::Member(Box::new(e), name);
            } else if self.is_p("!") && matches!(self.peek_at(1), Some(Tok::Ident(..))) {
                self.advance();
                let name = self.ident()?;
                e = Expr::Call(Box::new(e), vec![Arg { name: None, value: Some(Expr::Lit(Value::string(name.to_string()))) }]);
            } else if self.is_p("(") {
                let args = self.paren_args()?;
                e = Expr::Call(Box::new(e), args);
            } else {
                return Ok(e);
            }
        }
    }

    fn member_name(&mut self) -> PResult<Name> {
        self.ident()
    }

    fn primary(&mut self) -> PResult<Expr> {
        let tok = match self.peek() {
            Some(t) => t.clone(),
            None => return self.err("an expression was expected"),
        };
        match tok {
            Tok::Num(v) => {
                self.advance();
                Ok(Expr::Lit(v))
            }
            Tok::Str(s) => {
                self.advance();
                Ok(Expr::Lit(Value::Str(s)))
            }
            Tok::Date(d) => {
                self.advance();
                Ok(Expr::Lit(Value::Date(d)))
            }
            Tok::P("(") => {
                self.advance();
                let e = self.expr()?;
                self.expect_p(")")?;
                // Keeps `(a)(1)` and `(x).y` working through `postfix`.
                Ok(e)
            }
            Tok::P(".") => {
                self.advance();
                let name = self.member_name()?;
                Ok(Expr::WithMember(name))
            }
            Tok::Ident(n, _) => {
                self.advance();
                match &*n {
                    "new" => {
                        let mut class = self.ident()?;
                        while self.accept_p(".") {
                            class = self.ident()?;
                        }
                        Ok(Expr::New(class))
                    }
                    "typeof" => {
                        let e = self.binary(6)?;
                        self.expect_word("is")?;
                        let mut class = self.ident()?;
                        while self.accept_p(".") {
                            class = self.ident()?;
                        }
                        Ok(Expr::TypeOf(Box::new(e), class))
                    }
                    "true" => Ok(Expr::Lit(Value::Bool(true))),
                    "false" => Ok(Expr::Lit(Value::Bool(false))),
                    _ => Ok(Expr::Name(n)),
                }
            }
            _ => self.err("an expression was expected"),
        }
    }
}

/// Adds the arguments of a command-style call to the expression it names.
fn attach_args(target: Expr, args: Vec<Arg>) -> Expr {
    match target {
        // `obj.Method a, b` or `Name a, b`: the arguments belong to the last member or name.
        Expr::Member(..) | Expr::Name(_) | Expr::WithMember(_) => Expr::Call(Box::new(target), args),
        other => other,
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn parse(src: &str) -> Module {
        parse_module("M", src).unwrap()
    }

    #[test]
    fn procedures_and_parameters() {
        let m = parse(
            "Option Explicit\nPublic Function F(ByVal a As Double, Optional b = 2) As Double\n F = a + b\nEnd Function\nPrivate Sub S()\nEnd Sub\n",
        );
        assert_eq!(m.procs.len(), 2);
        assert_eq!(m.procs[0].params.len(), 2);
        assert!(m.procs[0].params[0].by_val);
        assert!(m.procs[0].params[1].optional);
        assert!(!m.procs[1].public);
    }

    #[test]
    fn blocks_and_single_line_if() {
        let m = parse(
            "Sub S()\nIf a > 1 Then b = 2 Else b = 3\nIf a Then\n x = 1\nElseIf b Then\n x = 2\nElse\n x = 3\nEnd If\nFor i = 1 To 10 Step 2\n s = s + i\nNext i\nDo\n x = x + 1\nLoop While x < 5\nEnd Sub",
        );
        assert_eq!(m.procs[0].body.len(), 4);
    }

    #[test]
    fn commands_without_parentheses_and_named_arguments() {
        let m = parse("Sub S()\nMsgBox \"x\", vbExclamation\nwb.SaveAs Filename:=p, FileFormat:=51\nws.Range(\"A1\").Value = 1\nEnd Sub");
        assert_eq!(m.procs[0].body.len(), 3);
        assert!(matches!(m.procs[0].body[1].kind, StmtKind::Call(Expr::Call(..))));
        assert!(matches!(m.procs[0].body[2].kind, StmtKind::Assign(..)));
    }

    #[test]
    fn precedence_follows_vba() {
        // -2 ^ 2 is -(2 ^ 2), and Not binds looser than comparison.
        let m = parse("Sub S()\nx = -2 ^ 2\ny = Not a = b\nEnd Sub");
        match &m.procs[0].body[0].kind {
            StmtKind::Assign(_, Expr::Unary(Un::Neg, inner)) => assert!(matches!(**inner, Expr::Binary(Bin::Pow, ..))),
            other => panic!("{other:?}"),
        }
        match &m.procs[0].body[1].kind {
            StmtKind::Assign(_, Expr::Unary(Un::Not, inner)) => assert!(matches!(**inner, Expr::Binary(Bin::Eq, ..))),
            other => panic!("{other:?}"),
        }
    }

    #[test]
    fn with_select_and_labels() {
        let m = parse(
            "Sub S()\nWith ws.Range(\"A1\")\n .Value = 1\nEnd With\nSelect Case n\nCase 1, 2\n x = 1\nCase Is > 5\n x = 2\nCase Else\n x = 3\nEnd Select\nOn Error GoTo Oops\nExit Sub\nOops:\nx = 0\nEnd Sub",
        );
        assert_eq!(m.procs[0].body.len(), 6);
        assert!(matches!(m.procs[0].body[3].kind, StmtKind::Exit(_)));
        assert!(matches!(m.procs[0].body[4].kind, StmtKind::Label(_)));
    }

    #[test]
    fn file_statements() {
        let m = parse("Sub S()\nOpen p For Output As #n\nPrint #n, \"a\" & vbTab & b\nClose #n\nEnd Sub");
        assert!(matches!(m.procs[0].body[0].kind, StmtKind::Open { .. }));
        assert!(matches!(m.procs[0].body[1].kind, StmtKind::Print { .. }));
    }

    #[test]
    fn deep_nesting_is_refused() {
        let src = format!("Sub S()\nx = {}1{}\nEnd Sub", "(".repeat(400), ")".repeat(400));
        assert!(parse_module("M", &src).is_err());
    }
}
