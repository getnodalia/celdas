//! The tree-walking interpreter: runs a [`Project`]'s procedures against a [`Host`].

use std::collections::HashMap;
use std::rc::Rc;

#[allow(clippy::wildcard_imports)]
use crate::ast::*;
use crate::host::{CallArg, Host};
use crate::native::{Collection, DebugObject, Dictionary, ErrObject};
use crate::ops;
use crate::project::Project;
use crate::value::{Obj, Result, VArray, Value, VbError};

/// Deepest nesting of procedure calls.
const MAX_CALL_DEPTH: u32 = 80;
/// Statements a run may execute before it is stopped as runaway.
pub const DEFAULT_MAX_STEPS: u64 = 400_000_000;
/// The error number of `End`: it unwinds every procedure and cannot be caught.
const END_SIGNAL: i32 = -1;

#[derive(Clone)]
struct Var {
    value: Value,
    ty: VType,
}

#[derive(Clone)]
enum OnErr {
    Off,
    ResumeNext,
    Goto(Name),
}

enum Flow {
    Next,
    Exit(ExitKind),
    Goto(Name),
}

struct Frame {
    vars: HashMap<Name, Var>,
    with: Vec<Value>,
    on_error: OnErr,
    module: usize,
}

impl Frame {
    fn new(module: usize) -> Frame {
        Frame { vars: HashMap::new(), with: Vec::new(), on_error: OnErr::Off, module }
    }
}

/// Where a run stopped on an error.
#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct ErrorSite {
    pub module: String,
    pub procedure: String,
    pub line: u32,
}

#[derive(Default)]
pub(crate) struct ErrState {
    pub number: i32,
    pub description: String,
}

pub struct Vm<'a> {
    pub(crate) host: &'a mut dyn Host,
    pub(crate) project: &'a Project,
    globals: HashMap<Name, Var>,
    pub(crate) err: ErrState,
    /// What `Debug.Print` wrote.
    pub debug: Vec<String>,
    steps: u64,
    max_steps: u64,
    depth: u32,
    /// File numbers currently open.
    pub(crate) files: Vec<i32>,
    pub(crate) rng: u64,
    pub(crate) last_rnd: f64,
    /// The procedure and line that raised the error that ended the run.
    site: Option<ErrorSite>,
    current: Vec<(usize, Name)>,
}

impl<'a> Vm<'a> {
    pub fn new(project: &'a Project, host: &'a mut dyn Host) -> Vm<'a> {
        Vm {
            host,
            project,
            globals: HashMap::new(),
            err: ErrState::default(),
            debug: Vec::new(),
            steps: 0,
            max_steps: DEFAULT_MAX_STEPS,
            depth: 0,
            files: Vec::new(),
            rng: 0x2545_F491_4F6C_DD1D,
            last_rnd: 0.5,
            site: None,
            current: Vec::new(),
        }
    }

    pub fn set_max_steps(&mut self, n: u64) {
        self.max_steps = n;
    }

    /// Where the error that ended the last run was raised.
    pub fn error_site(&self) -> Option<&ErrorSite> {
        self.site.as_ref()
    }

    pub fn steps(&self) -> u64 {
        self.steps
    }

    /// Runs the procedure `name` (`Proc` or `Module.Proc`) with `args`.
    pub fn run(&mut self, name: &str, args: Vec<Value>) -> Result<Value> {
        self.site = None;
        let project = self.project;
        let (module_hint, proc_name) = match name.rsplit_once('.') {
            Some((m, p)) => (Some(m), p),
            None => (None, name),
        };
        let wanted = proc_name.to_lowercase();
        let located = match module_hint {
            Some(m) => project
                .module_index(m)
                .and_then(|mi| project.find_proc(&wanted, Some(mi)))
                .filter(|(found, _)| project.module_index(module_hint.unwrap_or("")) == Some(*found)),
            None => project.find_proc(&wanted, None),
        };
        let Some((mi, p)) = located else {
            if let Some((module, e)) = project.syntax_errors().into_iter().next() {
                return Err(VbError::new(1004, format!("Compile error in {module}: {e}")));
            }
            return Err(VbError::not_defined(name));
        };
        self.init_globals()?;
        let argv = args.into_iter().map(CallArg::positional).collect();
        match self.call_proc(mi, p, argv) {
            Ok((v, _)) => Ok(v),
            Err(e) if e.number == END_SIGNAL => Ok(Value::Empty),
            Err(e) => Err(e),
        }
    }

    fn init_globals(&mut self) -> Result<()> {
        if !self.globals.is_empty() {
            return Ok(());
        }
        let project = self.project;
        for (mi, m) in project.modules.iter().enumerate() {
            let Ok(parsed) = &m.parsed else { continue };
            let mut f = Frame::new(mi);
            for d in &parsed.decls {
                match d {
                    Decl::Const(n, e) => {
                        let v = self.eval_r(&mut f, e)?;
                        self.globals.entry(n.clone()).or_insert(Var { value: v, ty: VType::Variant });
                    }
                    Decl::Var(item) => {
                        let var = self.declare(&mut f, item)?;
                        self.globals.entry(item.name.clone()).or_insert(var);
                    }
                }
            }
        }
        // A marker so a project without declarations is not initialised twice.
        self.globals.entry(Rc::from("")).or_insert(Var { value: Value::Empty, ty: VType::Variant });
        Ok(())
    }

    // ---- procedures ----

    fn call_proc(&mut self, module: usize, p: &'a Proc, args: Vec<CallArg>) -> Result<(Value, Vec<Value>)> {
        if self.depth >= MAX_CALL_DEPTH {
            return Err(VbError::new(28, "Out of stack space"));
        }
        let mut f = Frame::new(module);
        self.bind_params(&mut f, p, args)?;
        if p.kind == ProcKind::Function {
            f.vars.insert(p.name.clone(), Var { value: ops::default_value(&p.returns), ty: p.returns.clone() });
        }
        self.depth += 1;
        self.current.push((module, p.name.clone()));
        let flow = self.exec_block(&mut f, &p.body);
        self.current.pop();
        self.depth -= 1;
        flow?;
        let returned = if p.kind == ProcKind::Function { f.vars.get(&p.name).map(|v| v.value.clone()).unwrap_or(Value::Empty) } else { Value::Empty };
        let finals = p.params.iter().map(|q| f.vars.get(&q.name).map(|v| v.value.clone()).unwrap_or(Value::Empty)).collect();
        Ok((returned, finals))
    }

    fn bind_params(&mut self, f: &mut Frame, p: &Proc, args: Vec<CallArg>) -> Result<()> {
        let mut positional = args.iter().filter(|a| a.name.is_none()).map(|a| a.value.clone());
        for param in &p.params {
            if param.param_array {
                let rest: Vec<Value> = positional.by_ref().collect();
                f.vars.insert(param.name.clone(), Var { value: Value::array(VArray::from_vec(rest, 0)), ty: VType::Variant });
                continue;
            }
            let named = args.iter().find(|a| a.name.as_deref() == Some(&*param.name)).map(|a| a.value.clone());
            let given = named.or_else(|| positional.next()).filter(|v| !matches!(v, Value::Missing));
            let value = match given {
                Some(v) => {
                    let v = if matches!(param.ty, VType::Object(_)) { v } else { self.rvalue(v)? };
                    if matches!(v, Value::Array(_)) { v } else { ops::coerce(&param.ty, v)? }
                }
                None => match &param.default {
                    Some(e) => {
                        let d = self.eval_r(f, e)?;
                        ops::coerce(&param.ty, d)?
                    }
                    None if param.optional => Value::Missing,
                    None => return Err(VbError::new(449, "Argument not optional")),
                },
            };
            f.vars.insert(param.name.clone(), Var { value, ty: param.ty.clone() });
        }
        Ok(())
    }

    // ---- statements ----

    fn exec_block(&mut self, f: &mut Frame, stmts: &'a [Stmt]) -> Result<Flow> {
        let mut i = 0;
        while let Some(s) = stmts.get(i) {
            match self.exec_stmt(f, s)? {
                Flow::Next => i += 1,
                Flow::Goto(label) => match stmts.iter().position(|t| matches!(&t.kind, StmtKind::Label(l) if *l == label)) {
                    Some(at) => i = at + 1,
                    None => return Ok(Flow::Goto(label)),
                },
                other => return Ok(other),
            }
        }
        Ok(Flow::Next)
    }

    fn exec_stmt(&mut self, f: &mut Frame, s: &'a Stmt) -> Result<Flow> {
        self.steps += 1;
        if self.steps > self.max_steps {
            return Err(VbError::new(1004, "The macro ran for too long and was stopped"));
        }
        if self.steps.is_multiple_of(2048) {
            self.host.tick()?;
        }
        match self.exec_kind(f, s) {
            Ok(flow) => Ok(flow),
            Err(e) => self.on_error(f, s, e),
        }
    }

    fn on_error(&mut self, f: &mut Frame, s: &Stmt, e: VbError) -> Result<Flow> {
        if e.number == END_SIGNAL {
            return Err(e);
        }
        self.err = ErrState { number: e.number, description: e.description.clone() };
        match f.on_error.clone() {
            OnErr::ResumeNext => Ok(Flow::Next),
            OnErr::Goto(label) => {
                // While the handler runs, a new error is not caught by it.
                f.on_error = OnErr::Off;
                Ok(Flow::Goto(label))
            }
            OnErr::Off => {
                if self.site.is_none() {
                    let module = self.project.modules.get(f.module).map(|m| m.name.clone()).unwrap_or_default();
                    let procedure = self.current.last().map(|(_, n)| n.to_string()).unwrap_or_default();
                    self.site = Some(ErrorSite { module, procedure, line: s.line });
                }
                Err(e)
            }
        }
    }

    fn exec_kind(&mut self, f: &mut Frame, s: &'a Stmt) -> Result<Flow> {
        match &s.kind {
            StmtKind::If { .. }
            | StmtKind::For { .. }
            | StmtKind::ForEach { .. }
            | StmtKind::Do { .. }
            | StmtKind::While(..)
            | StmtKind::Select { .. }
            | StmtKind::With(..) => self.exec_control(f, &s.kind),
            StmtKind::Exit(k) => Ok(Flow::Exit(*k)),
            StmtKind::Goto(l) => Ok(Flow::Goto(l.clone())),
            StmtKind::Label(_) | StmtKind::Nop => Ok(Flow::Next),
            StmtKind::End => Err(VbError::new(END_SIGNAL, "End")),
            StmtKind::Resume(_) => Err(VbError::new(20, "Resume without error")),
            StmtKind::OnError(o) => {
                f.on_error = match o {
                    OnError::ResumeNext => OnErr::ResumeNext,
                    OnError::Goto0 => OnErr::Off,
                    OnError::Goto(l) => OnErr::Goto(l.clone()),
                };
                if matches!(o, OnError::Goto0) {
                    self.err = ErrState::default();
                }
                Ok(Flow::Next)
            }
            other => {
                self.exec_simple(f, other)?;
                Ok(Flow::Next)
            }
        }
    }

    fn exec_simple(&mut self, f: &mut Frame, kind: &'a StmtKind) -> Result<()> {
        match kind {
            StmtKind::Dim { items, .. } => {
                for item in items {
                    if let Some(existing) = f.vars.get_mut(&item.name) {
                        existing.ty = item.ty.clone();
                        continue;
                    }
                    let var = self.declare(f, item)?;
                    f.vars.insert(item.name.clone(), var);
                }
                Ok(())
            }
            StmtKind::Const(items) => {
                for (n, e) in items {
                    let v = self.eval_r(f, e)?;
                    f.vars.insert(n.clone(), Var { value: v, ty: VType::Variant });
                }
                Ok(())
            }
            StmtKind::ReDim { preserve, items } => items.iter().try_for_each(|i| self.redim(f, i, *preserve)),
            StmtKind::Assign(target, e) => {
                // A property of a host object takes the object itself (`series.Values = range`):
                // the host decides what that stands for. Anything else takes the default member.
                let v = if targets_member(target) { self.eval(f, e)? } else { self.eval_r(f, e)? };
                self.assign(f, target, v)
            }
            StmtKind::SetAssign(target, e) => {
                let v = self.eval(f, e)?;
                self.assign(f, target, v)
            }
            StmtKind::Call(e) => self.eval(f, e).map(|_| ()),
            StmtKind::Erase(targets) => {
                for t in targets {
                    if let Expr::Name(n) = t
                        && let Some(var) = self.var_mut(f, n)
                        && let Value::Array(a) = &var.value
                    {
                        let ty = var.ty.clone();
                        let mut a = a.borrow_mut();
                        for slot in &mut a.data {
                            *slot = ops::default_value(&ty);
                        }
                    }
                }
                Ok(())
            }
            StmtKind::Open { path, mode, number } => {
                let path = self.eval_r(f, path)?.to_text()?;
                let n = self.eval_r(f, number)?.to_i32()?;
                if self.files.contains(&n) {
                    return Err(VbError::new(55, "File already open"));
                }
                self.host.file_open(&path, *mode, n)?;
                self.files.push(n);
                Ok(())
            }
            StmtKind::Close(numbers) => {
                if numbers.is_empty() {
                    self.files.clear();
                    return self.host.file_close(None);
                }
                for e in numbers {
                    let n = self.eval_r(f, e)?.to_i32()?;
                    self.files.retain(|x| *x != n);
                    self.host.file_close(Some(n))?;
                }
                Ok(())
            }
            StmtKind::Print { number, items, newline, quoted } => self.exec_print(f, number, items, *newline, *quoted),
            StmtKind::LineInput { number, target } => {
                let n = self.eval_r(f, number)?.to_i32()?;
                match self.host.file_read_line(n)? {
                    Some(line) => self.assign(f, target, Value::string(line)),
                    None => Err(VbError::new(62, "Input past end of file")),
                }
            }
            _ => Ok(()),
        }
    }

    fn exec_print(&mut self, f: &mut Frame, number: &Expr, items: &[Expr], newline: bool, quoted: bool) -> Result<()> {
        let n = self.eval_r(f, number)?.to_i32()?;
        let mut text = String::new();
        for (i, e) in items.iter().enumerate() {
            let v = self.eval_r(f, e)?;
            if quoted && i > 0 {
                text.push(',');
            }
            text.push_str(&print_text(&v, quoted)?);
        }
        if newline {
            text.push_str("\r\n");
        }
        self.host.file_write(n, &text)
    }

    fn exec_control(&mut self, f: &mut Frame, kind: &'a StmtKind) -> Result<Flow> {
        match kind {
            StmtKind::If { branches, otherwise } => {
                for (cond, body) in branches {
                    if self.eval_r(f, cond)?.to_bool()? {
                        return self.exec_block(f, body);
                    }
                }
                match otherwise {
                    Some(body) => self.exec_block(f, body),
                    None => Ok(Flow::Next),
                }
            }
            StmtKind::For { var, from, to, step, body } => self.exec_for(f, var, from, to, step.as_ref(), body),
            StmtKind::ForEach { var, collection, body } => self.exec_for_each(f, var, collection, body),
            StmtKind::Do { cond, until, post, body } => self.exec_do(f, cond.as_ref(), *until, *post, body),
            StmtKind::While(cond, body) => {
                while self.eval_r(f, cond)?.to_bool()? {
                    self.steps += 1;
                    match self.exec_block(f, body)? {
                        Flow::Next => {}
                        Flow::Exit(ExitKind::Do) => break,
                        other => return Ok(other),
                    }
                }
                Ok(Flow::Next)
            }
            StmtKind::Select { subject, cases, otherwise } => {
                let v = self.eval_r(f, subject)?;
                for (tests, body) in cases {
                    for t in tests {
                        if self.case_matches(f, &v, t)? {
                            return self.exec_block(f, body);
                        }
                    }
                }
                match otherwise {
                    Some(body) => self.exec_block(f, body),
                    None => Ok(Flow::Next),
                }
            }
            StmtKind::With(target, body) => {
                let v = self.eval(f, target)?;
                f.with.push(v);
                let r = self.exec_block(f, body);
                f.with.pop();
                r
            }
            _ => Ok(Flow::Next),
        }
    }

    fn case_matches(&mut self, f: &mut Frame, v: &Value, t: &CaseTest) -> Result<bool> {
        Ok(match t {
            CaseTest::Value(e) => {
                let x = self.eval_r(f, e)?;
                ops::compare(v, &x)? == Some(std::cmp::Ordering::Equal)
            }
            CaseTest::Range(a, b) => {
                let (lo, hi) = (self.eval_r(f, a)?, self.eval_r(f, b)?);
                ops::compare(v, &lo)?.is_some_and(|o| o.is_ge()) && ops::compare(v, &hi)?.is_some_and(|o| o.is_le())
            }
            CaseTest::Is(op, e) => {
                let x = self.eval_r(f, e)?;
                matches!(ops::comparison(*op, v, &x)?, Value::Bool(true))
            }
        })
    }

    fn exec_for(&mut self, f: &mut Frame, var: &Expr, from: &Expr, to: &Expr, step: Option<&Expr>, body: &'a [Stmt]) -> Result<Flow> {
        let start = self.eval_r(f, from)?;
        let end = self.eval_r(f, to)?;
        let stp = match step {
            Some(e) => self.eval_r(f, e)?,
            None => Value::Int(1),
        };
        let integral = [&start, &end, &stp].iter().all(|v| matches!(v, Value::Int(_) | Value::Long(_) | Value::Byte(_)));
        let (mut i, limit, inc) = (start.to_f64()?, end.to_f64()?, stp.to_f64()?);
        if inc == 0.0 {
            return Err(VbError::invalid_call());
        }
        loop {
            if (inc > 0.0 && i > limit) || (inc < 0.0 && i < limit) {
                break;
            }
            self.steps += 1;
            let value = if integral { Value::Long(i as i32) } else { Value::Double(i) };
            self.assign(f, var, value)?;
            match self.exec_block(f, body)? {
                Flow::Next => {}
                Flow::Exit(ExitKind::For) => return Ok(Flow::Next),
                other => return Ok(other),
            }
            // The counter may have been changed by the body.
            i = self.eval_r(f, var)?.to_f64()? + inc;
            if self.steps > self.max_steps {
                return Err(VbError::new(1004, "The macro ran for too long and was stopped"));
            }
        }
        // After a completed loop the counter holds the first value past the limit.
        let past = if integral { Value::Long(i as i32) } else { Value::Double(i) };
        self.assign(f, var, past)?;
        Ok(Flow::Next)
    }

    fn exec_for_each(&mut self, f: &mut Frame, var: &Expr, collection: &Expr, body: &'a [Stmt]) -> Result<Flow> {
        let items = match self.eval(f, collection)? {
            Value::Array(a) => a.borrow().data.clone(),
            Value::Object(Some(o)) => self.enumerate(&o)?,
            _ => return Err(VbError::new(92, "For loop not initialized")),
        };
        for item in items {
            self.steps += 1;
            self.assign(f, var, item)?;
            match self.exec_block(f, body)? {
                Flow::Next => {}
                Flow::Exit(ExitKind::For) => return Ok(Flow::Next),
                other => return Ok(other),
            }
        }
        Ok(Flow::Next)
    }

    fn exec_do(&mut self, f: &mut Frame, cond: Option<&Expr>, until: bool, post: bool, body: &'a [Stmt]) -> Result<Flow> {
        loop {
            self.steps += 1;
            if self.steps > self.max_steps {
                return Err(VbError::new(1004, "The macro ran for too long and was stopped"));
            }
            if !post && !self.do_continues(f, cond, until)? {
                return Ok(Flow::Next);
            }
            match self.exec_block(f, body)? {
                Flow::Next => {}
                Flow::Exit(ExitKind::Do) => return Ok(Flow::Next),
                other => return Ok(other),
            }
            if post && !self.do_continues(f, cond, until)? {
                return Ok(Flow::Next);
            }
        }
    }

    fn do_continues(&mut self, f: &mut Frame, cond: Option<&Expr>, until: bool) -> Result<bool> {
        match cond {
            None => Ok(true),
            Some(c) => {
                let v = self.eval_r(f, c)?.to_bool()?;
                Ok(v != until)
            }
        }
    }

    // ---- variables ----

    fn declare(&mut self, f: &mut Frame, item: &DimItem) -> Result<Var> {
        if let Some(bounds) = &item.bounds {
            let mut dims = Vec::new();
            for (lo, hi) in bounds {
                let lo = match lo {
                    Some(e) => self.eval_r(f, e)?.to_i32()?,
                    None => 0,
                };
                dims.push((lo, self.eval_r(f, hi)?.to_i32()?));
            }
            let arr = VArray::new(dims, ops::default_value(&item.ty))?;
            return Ok(Var { value: Value::array(arr), ty: item.ty.clone() });
        }
        if item.dynamic {
            let arr = VArray { bounds: vec![(0, -1)], data: Vec::new() };
            return Ok(Var { value: Value::array(arr), ty: item.ty.clone() });
        }
        let value = if item.is_new {
            self.new_object(match &item.ty {
                VType::Object(c) => c,
                _ => return Err(VbError::type_mismatch()),
            })?
        } else {
            ops::default_value(&item.ty)
        };
        Ok(Var { value, ty: item.ty.clone() })
    }

    fn redim(&mut self, f: &mut Frame, item: &DimItem, preserve: bool) -> Result<()> {
        let Some(bounds) = &item.bounds else { return Err(VbError::invalid_call()) };
        let mut dims = Vec::new();
        for (lo, hi) in bounds {
            let lo = match lo {
                Some(e) => self.eval_r(f, e)?.to_i32()?,
                None => 0,
            };
            dims.push((lo, self.eval_r(f, hi)?.to_i32()?));
        }
        let ty = match self.var_mut(f, &item.name) {
            Some(v) if !matches!(item.ty, VType::Variant) || matches!(v.ty, VType::Variant) => {
                if matches!(item.ty, VType::Variant) {
                    v.ty.clone()
                } else {
                    item.ty.clone()
                }
            }
            Some(v) => v.ty.clone(),
            None => item.ty.clone(),
        };
        let mut arr = VArray::new(dims, ops::default_value(&ty))?;
        if preserve && let Some(old) = self.var_mut(f, &item.name).and_then(|v| if let Value::Array(a) = &v.value { Some(a.clone()) } else { None }) {
            copy_overlap(&old.borrow(), &mut arr);
        }
        let value = Value::array(arr);
        match self.var_mut(f, &item.name) {
            Some(v) => {
                v.value = value;
                v.ty = ty;
            }
            None => {
                f.vars.insert(item.name.clone(), Var { value, ty });
            }
        }
        Ok(())
    }

    fn var_mut<'x>(&'x mut self, f: &'x mut Frame, n: &Name) -> Option<&'x mut Var> {
        if f.vars.contains_key(n) { f.vars.get_mut(n) } else { self.globals.get_mut(n) }
    }

    fn var_exists(&self, f: &Frame, n: &Name) -> bool {
        f.vars.contains_key(n) || self.globals.contains_key(n)
    }

    fn assign_var(&mut self, f: &mut Frame, n: &Name, value: Value) -> Result<()> {
        let value = match value {
            // Arrays are copied on assignment.
            Value::Array(a) => Value::array(a.borrow().clone()),
            other => other,
        };
        match self.var_mut(f, n) {
            Some(v) => {
                if matches!(value, Value::Array(_)) {
                    v.value = value;
                } else {
                    v.value = ops::coerce(&v.ty, value)?;
                }
            }
            None => {
                f.vars.insert(n.clone(), Var { value, ty: VType::Variant });
            }
        }
        Ok(())
    }

    fn assign(&mut self, f: &mut Frame, target: &Expr, value: Value) -> Result<()> {
        match target {
            Expr::Name(n) => self.assign_var(f, n, value),
            Expr::Call(base, args) => self.assign_call(f, base, args, value),
            Expr::Member(base, name) => {
                let obj = self.eval_object(f, base)?;
                self.assign_member(&obj, name, &[], value)
            }
            Expr::WithMember(name) => {
                let obj = self.with_object(f)?;
                self.assign_member(&obj, name, &[], value)
            }
            _ => Err(VbError::new(451, "Property let procedure not defined and property get procedure did not return an object")),
        }
    }

    /// Sets a property. The interpreter's own objects take the value of an object assigned to
    /// them; a host object gets it as it is.
    fn assign_member(&mut self, obj: &Obj, name: &str, args: &[CallArg], value: Value) -> Result<()> {
        let native = obj.downcast::<ErrObject>().is_some()
            || obj.downcast::<DebugObject>().is_some()
            || obj.downcast::<Collection>().is_some()
            || obj.downcast::<Dictionary>().is_some();
        let value = if native { self.rvalue(value)? } else { value };
        self.member(obj, name, args, Some(value)).map(|_| ())
    }

    fn assign_call(&mut self, f: &mut Frame, base: &Expr, args: &[Arg], value: Value) -> Result<()> {
        let argv = self.eval_args(f, args)?;
        match base {
            Expr::Name(n) if self.var_exists(f, n) => {
                let held = self.var_mut(f, n).map(|var| (var.value.clone(), var.ty.clone()));
                let Some((held, ty)) = held else { return Err(VbError::subscript()) };
                match &held {
                    Value::Array(a) => {
                        let idx: Vec<i32> = argv.iter().map(|a| self.rvalue(a.value.clone()).and_then(|v| v.to_i32())).collect::<Result<_>>()?;
                        let v = ops::coerce(&ty, value)?;
                        a.borrow_mut().set(&idx, v)
                    }
                    Value::Object(Some(o)) => {
                        let o = o.clone();
                        self.member(&o, "", &argv, Some(value)).map(|_| ())
                    }
                    _ => Err(VbError::subscript()),
                }
            }
            Expr::Name(n) => match self.host.global(n, &argv, Some(value))? {
                Some(_) => Ok(()),
                None => Err(VbError::not_defined(n)),
            },
            Expr::Member(b, name) => {
                let obj = self.eval_object(f, b)?;
                self.assign_member(&obj, name, &argv, value)
            }
            Expr::WithMember(name) => {
                let obj = self.with_object(f)?;
                self.assign_member(&obj, name, &argv, value)
            }
            _ => Err(VbError::new(451, "Property let procedure not defined and property get procedure did not return an object")),
        }
    }

    fn with_object(&mut self, f: &Frame) -> Result<Obj> {
        match f.with.last() {
            Some(Value::Object(Some(o))) => Ok(o.clone()),
            Some(Value::Object(None)) | None => Err(VbError::not_set()),
            Some(_) => Err(VbError::object_required()),
        }
    }

    // ---- expressions ----

    /// The value of `e` for use as an operand: an object stands for its default member.
    fn eval_r(&mut self, f: &mut Frame, e: &Expr) -> Result<Value> {
        let v = self.eval(f, e)?;
        self.rvalue(v)
    }

    pub(crate) fn rvalue(&mut self, v: Value) -> Result<Value> {
        match v {
            Value::Object(Some(o)) => {
                if o.downcast::<Collection>().is_some() || o.downcast::<Dictionary>().is_some() {
                    return Ok(Value::Object(Some(o)));
                }
                let d = self.member(&o, "", &[], None)?;
                // A default member that is itself an object (a Range's Item) is not unwrapped twice.
                Ok(d)
            }
            Value::Object(None) => Err(VbError::not_set()),
            other => Ok(other),
        }
    }

    fn eval_object(&mut self, f: &mut Frame, e: &Expr) -> Result<Obj> {
        match self.eval(f, e)? {
            Value::Object(Some(o)) => Ok(o),
            Value::Object(None) => Err(VbError::not_set()),
            _ => Err(VbError::object_required()),
        }
    }

    fn eval(&mut self, f: &mut Frame, e: &Expr) -> Result<Value> {
        match e {
            Expr::Lit(v) => Ok(v.clone()),
            Expr::Name(n) => self.eval_name(f, n),
            Expr::WithMember(n) => {
                let obj = self.with_object(f)?;
                self.member(&obj, n, &[], None)
            }
            Expr::Member(base, n) => self.eval_member(f, base, n, Vec::new()),
            Expr::Call(base, args) => self.eval_call(f, base, args),
            Expr::Unary(op, x) => {
                let v = self.eval_r(f, x)?;
                ops::unary(*op, &v)
            }
            Expr::Binary(op, a, b) => self.eval_binary(f, *op, a, b),
            Expr::New(c) => self.new_object(c),
            Expr::TypeOf(x, class) => {
                let v = self.eval(f, x)?;
                Ok(Value::Bool(match v {
                    Value::Object(Some(o)) => o.class().eq_ignore_ascii_case(class) || &**class == "object",
                    _ => false,
                }))
            }
        }
    }

    fn eval_binary(&mut self, f: &mut Frame, op: Bin, a: &Expr, b: &Expr) -> Result<Value> {
        if op == Bin::Is {
            let (x, y) = (self.eval(f, a)?, self.eval(f, b)?);
            return Ok(Value::Bool(match (&x, &y) {
                (Value::Object(p), Value::Object(q)) => match (p, q) {
                    (None, None) => true,
                    (Some(p), Some(q)) => p.same(q),
                    _ => false,
                },
                _ => return Err(VbError::object_required()),
            }));
        }
        let x = self.eval_r(f, a)?;
        let y = self.eval_r(f, b)?;
        match op {
            Bin::Add | Bin::Sub | Bin::Mul | Bin::Div | Bin::IntDiv | Bin::Mod | Bin::Pow => ops::arith(op, &x, &y),
            Bin::Concat => ops::concat(&x, &y),
            Bin::Eq | Bin::Ne | Bin::Lt | Bin::Gt | Bin::Le | Bin::Ge => ops::comparison(op, &x, &y),
            Bin::Like => Ok(Value::Bool(ops::like(&x.to_text()?, &y.to_text()?))),
            _ => ops::logic(op, &x, &y),
        }
    }

    fn eval_name(&mut self, f: &mut Frame, n: &Name) -> Result<Value> {
        if let Some(v) = f.vars.get(n).or_else(|| self.globals.get(n)) {
            return Ok(v.value.clone());
        }
        self.eval_free_name(f, n, &[])
    }

    /// A name that is not a variable: a constant, a procedure, a built-in or a host name.
    fn eval_free_name(&mut self, f: &mut Frame, n: &Name, args: &[CallArg]) -> Result<Value> {
        match &**n {
            "nothing" => return Ok(Value::nothing()),
            "empty" => return Ok(Value::Empty),
            "null" => return Ok(Value::Null),
            "err" => return Ok(Value::object(ErrObject)),
            "debug" => return Ok(Value::object(DebugObject)),
            _ => {}
        }
        if args.is_empty()
            && let Some(v) = crate::builtins::vb_constant(n)
        {
            return Ok(v);
        }
        let project = self.project;
        if let Some((mi, p)) = project.find_proc(n, Some(f.module)) {
            return self.call_proc(mi, p, args.to_vec()).map(|(v, _)| v);
        }
        if let Some(r) = self.builtin(n, args) {
            return r;
        }
        match self.host.global(n, args, None)? {
            Some(v) => Ok(v),
            None => Err(VbError::not_defined(n)),
        }
    }

    fn eval_args(&mut self, f: &mut Frame, args: &[Arg]) -> Result<Vec<CallArg>> {
        let mut out = Vec::with_capacity(args.len());
        for a in args {
            let value = match &a.value {
                Some(e) => self.eval(f, e)?,
                None => Value::Missing,
            };
            out.push(CallArg { name: a.name.as_ref().map(|n| n.to_string()), value });
        }
        Ok(out)
    }

    fn eval_call(&mut self, f: &mut Frame, base: &Expr, args: &[Arg]) -> Result<Value> {
        match base {
            Expr::Name(n) => self.call_name(f, n, args),
            Expr::Member(b, name) => {
                let argv = self.eval_args(f, args)?;
                self.eval_member(f, b, name, argv)
            }
            Expr::WithMember(name) => {
                let argv = self.eval_args(f, args)?;
                let obj = self.with_object(f)?;
                self.member(&obj, name, &argv, None)
            }
            other => {
                let argv = self.eval_args(f, args)?;
                let v = self.eval(f, other)?;
                self.index_value(v, &argv)
            }
        }
    }

    fn call_name(&mut self, f: &mut Frame, n: &Name, args: &[Arg]) -> Result<Value> {
        // A variable that is not an array or an object cannot be indexed: a function calling
        // itself (`Fact = n * Fact(n - 1)`) names its own return variable.
        let held = f.vars.get(n).or_else(|| self.globals.get(n)).map(|v| v.value.clone()).filter(|v| matches!(v, Value::Array(_) | Value::Object(_)));
        if let Some(v) = held {
            let argv = self.eval_args(f, args)?;
            return self.index_value(v, &argv);
        }
        let project = self.project;
        if let Some((mi, p)) = project.find_proc(n, Some(f.module)) {
            let argv = self.eval_args(f, args)?;
            let (ret, finals) = self.call_proc(mi, p, argv)?;
            // By-reference arguments that were plain variables take the callee's final value.
            for (i, (a, param)) in args.iter().zip(&p.params).enumerate() {
                if let (Some(Expr::Name(var)), false) = (&a.value, param.by_val)
                    && let Some(v) = finals.get(i)
                    && !matches!(param.ty, VType::Object(_))
                    && self.var_exists(f, var)
                {
                    self.assign_var(f, var, v.clone())?;
                }
            }
            return Ok(ret);
        }
        let argv = self.eval_args(f, args)?;
        self.eval_free_name(f, n, &argv)
    }

    /// `value(args)` for a value that is not a procedure: an array element or an object's
    /// default member.
    fn index_value(&mut self, v: Value, args: &[CallArg]) -> Result<Value> {
        match v {
            Value::Array(a) => {
                let idx: Vec<i32> = args.iter().map(|x| self.rvalue(x.value.clone()).and_then(|v| v.to_i32())).collect::<Result<_>>()?;
                a.borrow().get(&idx)
            }
            Value::Object(Some(o)) => self.member(&o, "", args, None),
            Value::Object(None) => Err(VbError::not_set()),
            _ => Err(VbError::new(13, "Type mismatch")),
        }
    }

    fn eval_member(&mut self, f: &mut Frame, base: &Expr, name: &Name, args: Vec<CallArg>) -> Result<Value> {
        // `Module.Procedure` and `VBA.Function`.
        if let Expr::Name(m) = base
            && !self.var_exists(f, m)
        {
            let project = self.project;
            if let Some(mi) = project.module_index(m)
                && let Some((found, p)) = project.find_proc(name, Some(mi))
                && found == mi
            {
                return self.call_proc(mi, p, args).map(|(v, _)| v);
            }
            if matches!(&**m, "vba" | "vbmath" | "strings" | "conversion" | "interaction" | "information" | "datetime" | "filesystem" | "math")
                && let Some(r) = self.builtin(name, &args)
            {
                return r;
            }
        }
        let obj = match self.eval(f, base)? {
            Value::Object(Some(o)) => o,
            Value::Object(None) => return Err(VbError::not_set()),
            _ => return Err(VbError::object_required()),
        };
        self.member(&obj, name, &args, None)
    }

    /// A member of an object: VBA's own objects here, everything else the host's.
    pub(crate) fn member(&mut self, obj: &Obj, name: &str, args: &[CallArg], assign: Option<Value>) -> Result<Value> {
        if obj.downcast::<ErrObject>().is_some() {
            return self.err_member(name, args, assign);
        }
        if obj.downcast::<DebugObject>().is_some() {
            return self.debug_member(name, args);
        }
        if let Some(c) = obj.downcast::<Collection>() {
            return c.member(name, args);
        }
        if let Some(d) = obj.downcast::<Dictionary>() {
            return d.member(name, args, assign);
        }
        self.host.member(obj, name, args, assign)
    }

    fn enumerate(&mut self, obj: &Obj) -> Result<Vec<Value>> {
        if let Some(c) = obj.downcast::<Collection>() {
            return Ok(c.items());
        }
        if let Some(d) = obj.downcast::<Dictionary>() {
            return Ok(d.keys());
        }
        self.host.enumerate(obj)
    }

    fn err_member(&mut self, name: &str, args: &[CallArg], assign: Option<Value>) -> Result<Value> {
        match name {
            "" | "number" => {
                if let Some(v) = assign {
                    self.err.number = v.to_i32()?;
                    return Ok(Value::Empty);
                }
                Ok(Value::Long(self.err.number))
            }
            "description" => {
                if let Some(v) = assign {
                    self.err.description = v.to_text()?;
                    return Ok(Value::Empty);
                }
                Ok(Value::string(self.err.description.clone()))
            }
            "source" => Ok(Value::str("VBAProject")),
            "clear" => {
                self.err = ErrState::default();
                Ok(Value::Empty)
            }
            "raise" => {
                let number = crate::host::arg(args, 0, "number").map(|v| v.to_i32()).transpose()?.unwrap_or(0);
                let description = match crate::host::arg(args, 2, "description") {
                    Some(d) => d.to_text()?,
                    None => "Application-defined or object-defined error".to_string(),
                };
                Err(VbError::new(if number == 0 { 5 } else { number }, description))
            }
            other => Err(VbError::no_member(other)),
        }
    }

    fn debug_member(&mut self, name: &str, args: &[CallArg]) -> Result<Value> {
        match name {
            "print" => {
                let mut parts = Vec::new();
                for a in args {
                    let v = self.rvalue(a.value.clone())?;
                    parts.push(v.to_text()?);
                }
                self.debug.push(parts.join("\t"));
                Ok(Value::Empty)
            }
            "assert" => Ok(Value::Empty),
            other => Err(VbError::no_member(other)),
        }
    }

    pub(crate) fn new_object(&mut self, class: &str) -> Result<Value> {
        match class {
            "collection" => Ok(Value::object(Collection::default())),
            "dictionary" => Ok(Value::object(Dictionary::default())),
            other => self.host.create_object(other),
        }
    }
}

/// The text `Print #`/`Write #` writes for a value.
fn print_text(v: &Value, quoted: bool) -> Result<String> {
    Ok(match v {
        Value::Str(s) => {
            if quoted {
                format!("\"{s}\"")
            } else {
                s.to_string()
            }
        }
        Value::Bool(b) => if *b { "True" } else { "False" }.to_string(),
        Value::Null => if quoted { "#NULL#" } else { "Null" }.to_string(),
        Value::Empty | Value::Missing => String::new(),
        other => {
            let t = other.to_text()?;
            if quoted || matches!(other, Value::Date(_)) {
                t
            } else if other.as_number().is_some_and(|n| n >= 0.0) {
                format!(" {t} ")
            } else {
                format!("{t} ")
            }
        }
    })
}

/// Copies the elements two arrays share (same indexes) from `old` into `new`.
fn copy_overlap(old: &VArray, new: &mut VArray) {
    if old.bounds.len() != new.bounds.len() {
        return;
    }
    let mut idx: Vec<i32> = new.bounds.iter().map(|b| b.0).collect();
    loop {
        if let Ok(v) = old.get(&idx) {
            let _ = new.set(&idx, v);
        }
        // Advance the odometer, last dimension fastest.
        let mut d = idx.len();
        loop {
            if d == 0 {
                return;
            }
            d -= 1;
            let Some(&(_, hi)) = new.bounds.get(d) else { return };
            let Some(slot) = idx.get_mut(d) else { return };
            if *slot < hi {
                *slot += 1;
                break;
            }
            *slot = new.bounds.get(d).map_or(0, |b| b.0);
        }
    }
}

/// Whether an assignment target is a property of an object (`a.b = …`, `.b = …`, `a.b(1) = …`).
fn targets_member(target: &Expr) -> bool {
    match target {
        Expr::Member(..) | Expr::WithMember(_) => true,
        Expr::Call(base, _) => matches!(**base, Expr::Member(..) | Expr::WithMember(_)),
        _ => false,
    }
}
