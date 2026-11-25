//! Stepping through a macro. The interpreter keeps no frames it could freeze, so a pause ends the
//! run and the next step runs it again from the start, answering each earlier pause the way the
//! person did. The run is the same every time, so it stops at the same place: the caller restores the
//! workbook it started from before each replay, and gets a [`Pause`] describing where the run is.

use std::collections::BTreeSet;

use crate::ast::{Stmt, StmtKind, VType};
use crate::interp::{Frame, Vm};
use crate::value::{Result, Value, VbError};

/// The error number that unwinds a run to its pause. Like `End` it cannot be caught.
pub(crate) const PAUSE_SIGNAL: i32 = -2;
/// Most text of one value a pause shows.
const MAX_SHOWN: usize = 200;
/// Most immediate statements one process compiles (their syntax trees are kept for the process's life).
const MAX_SNIPPETS: usize = 10_000;

/// How to go on from a pause.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum Step {
    /// To the next statement, entering the procedures it calls.
    Into,
    /// To the next statement of this procedure; calls run through.
    Over,
    /// Until this procedure returns.
    Out,
    /// Until a breakpoint or `Stop`.
    Continue,
}

/// A breakpoint: a module (any case) and a 1-based line of its source.
pub type Breakpoint = (String, u32);

/// One answer to a pause.
#[derive(Clone, Debug, PartialEq, Eq)]
pub enum Command {
    /// Go on this way, with these breakpoints from now on.
    Go(Step, BTreeSet<Breakpoint>),
    /// Run a statement (or, after `?`, show an expression) in the paused procedure.
    Immediate(String),
}

/// What to do: the answers given so far, in order. The run pauses at the first pause no answer covers;
/// with no answers, at its first statement.
#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct Plan {
    pub script: Vec<Command>,
    /// Expressions to evaluate in the paused procedure.
    pub watches: Vec<String>,
}

/// A variable at a pause.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Variable {
    pub name: String,
    pub ty: String,
    pub value: String,
}

/// A procedure on the call stack.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct StackFrame {
    pub module: String,
    pub procedure: String,
    /// The line it is at (1-based, of the module's source).
    pub line: u32,
    pub locals: Vec<Variable>,
}

/// Why the run stopped here.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum Reason {
    Step,
    Breakpoint,
    Stop,
}

/// Where a paused run is.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Pause {
    pub reason: Reason,
    /// The innermost procedure first.
    pub frames: Vec<StackFrame>,
    /// Module-level variables.
    pub globals: Vec<Variable>,
    /// Each watch with its value, or the error that evaluating it gave.
    pub watches: Vec<(String, std::result::Result<String, String>)>,
    /// What `Debug.Print` has written so far.
    pub output: Vec<String>,
    /// The results of the immediate statements run at this pause, in order.
    pub immediate: Vec<String>,
}

/// How a debugged run went.
#[derive(Debug)]
pub enum Outcome {
    Paused(Box<Pause>),
    Finished(Result<Value>),
}

#[derive(Clone, Copy, Debug)]
enum Mode {
    Into,
    Over(u32),
    Out(u32),
    Continue,
}

/// What the interpreter keeps while it is debugged.
pub(crate) struct State {
    plan: Plan,
    next: usize,
    mode: Mode,
    breakpoints: BTreeSet<(String, u32)>,
    last: Option<(usize, u32, u32)>,
    /// The line each active procedure is at.
    pub(crate) lines: Vec<u32>,
    pending: Option<Pause>,
    immediate: Vec<String>,
}

impl State {
    fn new(plan: Plan) -> State {
        State { plan, next: 0, mode: Mode::Into, breakpoints: BTreeSet::new(), last: None, lines: Vec::new(), pending: None, immediate: Vec::new() }
    }
}

/// How a value is shown.
pub(crate) fn show(v: &Value) -> String {
    let text = match v {
        Value::Str(s) => format!("\"{s}\""),
        Value::Empty | Value::Missing => "Empty".into(),
        Value::Null => "Null".into(),
        Value::Object(None) => "Nothing".into(),
        Value::Object(Some(o)) => format!("<{}>", o.class()),
        Value::Error(n) => format!("Error {n}"),
        Value::Array(a) => {
            let dims: Vec<String> = a.borrow().bounds.iter().map(|(lo, hi)| format!("{lo} To {hi}")).collect();
            format!("Array({})", dims.join(", "))
        }
        other => other.to_text().unwrap_or_default(),
    };
    if text.chars().count() > MAX_SHOWN { format!("{}…", text.chars().take(MAX_SHOWN).collect::<String>()) } else { text }
}

fn type_label(ty: &VType, v: &Value) -> String {
    match ty {
        VType::Variant => format!("Variant/{}", v.type_name()),
        VType::Boolean => "Boolean".into(),
        VType::Byte => "Byte".into(),
        VType::Integer => "Integer".into(),
        VType::Long => "Long".into(),
        VType::Single => "Single".into(),
        VType::Double => "Double".into(),
        VType::Currency => "Currency".into(),
        VType::Date => "Date".into(),
        VType::String => "String".into(),
        VType::Object(n) => v.type_name().chars().next().map_or(n.to_string(), |_| v.type_name()),
    }
}

/// A syntax tree kept for the life of the process: the interpreter borrows its tree for as long as
/// its project, and an immediate statement has no project to live in.
fn keep<T>(v: T) -> &'static T {
    Box::leak(Box::new(v))
}

impl<'a> Vm<'a> {
    /// Runs `name` under the debugger, pausing as `plan` says.
    pub fn run_debug(&mut self, name: &str, args: Vec<Value>, plan: Plan) -> Outcome {
        self.dbg = Some(Box::new(State::new(plan)));
        let result = self.run(name, args);
        let state = self.dbg.take();
        match (result, state.and_then(|s| s.pending)) {
            (Err(e), Some(pause)) if e.number == PAUSE_SIGNAL => Outcome::Paused(Box::new(pause)),
            (r, _) => Outcome::Finished(r),
        }
    }

    pub(crate) fn variables(f: &Frame) -> Vec<Variable> {
        let mut out: Vec<Variable> =
            f.vars.iter().map(|(n, v)| Variable { name: n.to_string(), ty: type_label(&v.ty, &v.value), value: show(&v.value) }).collect();
        out.sort_by(|a, b| a.name.cmp(&b.name));
        out
    }

    /// Called before each statement of a debugged run; `Some` pauses it.
    pub(crate) fn debug_hook(&mut self, f: &mut Frame, s: &'a Stmt) -> Option<VbError> {
        let depth = self.depth;
        let module_name = self.project.modules.get(f.module).map(|m| m.name.to_lowercase()).unwrap_or_default();
        let st = self.dbg.as_mut()?;
        if let Some(l) = st.lines.last_mut() {
            *l = s.line;
        }
        if matches!(s.kind, StmtKind::Label(_) | StmtKind::Nop | StmtKind::Dim { .. } | StmtKind::Const(_)) {
            return None;
        }
        let here = (f.module, s.line, depth);
        let moved = st.last != Some(here);
        st.last = Some(here);
        let stop = matches!(s.kind, StmtKind::Stop);
        if !moved && !stop {
            return None;
        }
        let on_breakpoint = st.breakpoints.contains(&(module_name, s.line));
        let by_mode = match st.mode {
            Mode::Into => true,
            Mode::Over(d) => depth <= d,
            Mode::Out(d) => depth < d,
            Mode::Continue => false,
        };
        if !(on_breakpoint || stop || by_mode) {
            return None;
        }
        // A pause the earlier answers cover: answer it again, the way it was answered.
        loop {
            let st = self.dbg.as_mut()?;
            match st.plan.script.get(st.next).cloned() {
                Some(Command::Go(step, breakpoints)) => {
                    st.next += 1;
                    st.immediate.clear();
                    st.breakpoints = breakpoints.into_iter().map(|(m, l)| (m.to_lowercase(), l)).collect();
                    st.mode = match step {
                        Step::Into => Mode::Into,
                        Step::Over => Mode::Over(depth),
                        Step::Out => Mode::Out(depth),
                        Step::Continue => Mode::Continue,
                    };
                    return None;
                }
                Some(Command::Immediate(text)) => {
                    st.next += 1;
                    let shown = self.immediate(f, &text);
                    if let Some(st) = self.dbg.as_mut() {
                        st.immediate.push(shown);
                    }
                }
                None => break,
            }
        }
        let reason = if stop {
            Reason::Stop
        } else if on_breakpoint {
            Reason::Breakpoint
        } else {
            Reason::Step
        };
        let watches: Vec<String> = self.dbg.as_ref().map(|st| st.plan.watches.clone()).unwrap_or_default();
        let watches = watches.into_iter().map(|w| self.watch(f, &w)).collect();
        let mut globals = self.global_variables();
        globals.retain(|v| !v.name.is_empty());
        let immediate = self.dbg.as_ref().map(|st| st.immediate.clone()).unwrap_or_default();
        let pause = Pause { reason, frames: Vec::new(), globals, watches, output: self.debug.clone(), immediate };
        if let Some(st) = self.dbg.as_mut() {
            st.pending = Some(pause);
        }
        Some(VbError::new(PAUSE_SIGNAL, "Paused"))
    }

    /// Adds the procedure being left to the pause that is unwinding through it.
    pub(crate) fn debug_frame(&mut self, f: &Frame, procedure: &str) {
        let module = self.project.modules.get(f.module).map(|m| m.name.clone()).unwrap_or_default();
        let Some(st) = self.dbg.as_mut() else { return };
        let line = st.lines.last().copied().unwrap_or(0);
        if let Some(p) = st.pending.as_mut() {
            p.frames.push(StackFrame { module, procedure: procedure.to_string(), line, locals: Vm::variables(f) });
        }
    }

    fn global_variables(&self) -> Vec<Variable> {
        let mut out: Vec<Variable> =
            self.globals_iter().map(|(n, v)| Variable { name: n.to_string(), ty: type_label(&v.ty, &v.value), value: show(&v.value) }).collect();
        out.sort_by(|a, b| a.name.cmp(&b.name));
        out
    }

    fn watch(&mut self, f: &mut Frame, text: &str) -> (String, std::result::Result<String, String>) {
        let shown = match crate::parser::parse_expression(text) {
            Err(e) => Err(e.message),
            Ok(expr) => {
                let expr: &'static crate::ast::Expr = keep(expr);
                match self.eval_r(f, expr) {
                    Ok(v) => Ok(show(&v)),
                    Err(e) => Err(e.description),
                }
            }
        };
        (text.to_string(), shown)
    }

    /// Runs one line of the Immediate window in the paused procedure and says what it printed.
    fn immediate(&mut self, f: &mut Frame, text: &str) -> String {
        if self.snippets >= MAX_SNIPPETS {
            return "Too many statements were run in this session".into();
        }
        self.snippets += 1;
        let trimmed = text.trim();
        if let Some(expr) = trimmed.strip_prefix('?') {
            return match crate::parser::parse_expression(expr) {
                Err(e) => e.message,
                Ok(e) => match self.eval_r(f, keep(e)) {
                    Ok(v) => v.to_text().unwrap_or_else(|_| show(&v)),
                    Err(e) => format!("Error {}: {}", e.number, e.description),
                },
            };
        }
        match crate::parser::parse_snippet(trimmed) {
            Err(e) => e.message,
            Ok(stmts) => {
                let stmts: &'static Vec<Stmt> = keep(stmts);
                let before = self.debug.len();
                let paused = self.dbg.take();
                let r = self.exec_block(f, stmts);
                self.dbg = paused;
                match r {
                    Ok(_) => self.debug.get(before..).map(|d| d.join("\n")).unwrap_or_default(),
                    Err(e) => format!("Error {}: {}", e.number, e.description),
                }
            }
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::host::NoHost;
    use crate::project::{ModuleKind, Project, SourceModule};

    const SRC: &str = "Dim Total As Long\nFunction Twice(n As Long) As Long\n    Twice = n * 2\nEnd Function\nSub Main()\n    Dim i As Long\n    Dim s As String\n    s = \"hi\"\n    For i = 1 To 3\n        Total = Total + Twice(i)\n    Next i\n    Debug.Print Total\nEnd Sub\n";

    fn project() -> Project {
        Project::from_sources(vec![SourceModule { name: "Module1".into(), kind: ModuleKind::Standard, source: SRC.into() }])
    }

    fn go(step: Step, bps: &[u32]) -> Command {
        Command::Go(step, bps.iter().map(|l| ("module1".to_string(), *l)).collect())
    }

    fn run(script: Vec<Command>, watches: &[&str]) -> Outcome {
        let p = project();
        let mut host = NoHost;
        let mut vm = Vm::new(&p, &mut host);
        vm.run_debug("Main", Vec::new(), Plan { script, watches: watches.iter().map(|w| w.to_string()).collect() })
    }

    fn paused(o: Outcome) -> Pause {
        match o {
            Outcome::Paused(p) => *p,
            Outcome::Finished(r) => panic!("finished: {r:?}"),
        }
    }

    #[test]
    fn a_run_pauses_at_its_first_statement_and_steps_into_calls() {
        let p = paused(run(vec![], &[]));
        assert_eq!((p.frames.len(), p.frames[0].line), (1, 8));
        // Statement `s = "hi"`, then `For`, then the line of the body, then into Twice.
        let p = paused(run(vec![go(Step::Into, &[]), go(Step::Into, &[]), go(Step::Into, &[])], &[]));
        assert_eq!((p.frames.len(), p.frames[0].procedure.as_str(), p.frames[0].line), (2, "Twice", 3));
        assert_eq!(p.frames[1].procedure, "Main");
        assert_eq!(p.frames[1].line, 10);
        let n = p.frames[0].locals.iter().find(|v| v.name == "n").unwrap();
        assert_eq!((n.value.as_str(), n.ty.as_str()), ("1", "Long"));
    }

    #[test]
    fn over_runs_a_call_through_and_out_returns_from_it() {
        let over = paused(run(vec![go(Step::Over, &[]), go(Step::Over, &[]), go(Step::Over, &[]), go(Step::Over, &[])], &[]));
        assert_eq!((over.frames.len(), over.frames[0].line), (1, 10));
        let i = over.frames[0].locals.iter().find(|v| v.name == "i").map(|v| v.value.clone());
        assert_eq!(i.as_deref(), Some("3"));
        let script = vec![go(Step::Into, &[]), go(Step::Into, &[]), go(Step::Into, &[]), go(Step::Out, &[])];
        let out = paused(run(script, &[]));
        assert_eq!((out.frames.len(), out.frames[0].procedure.as_str(), out.frames[0].line), (1, "Main", 10));
    }

    #[test]
    fn continue_stops_at_a_breakpoint_each_time_and_then_finishes() {
        let p = paused(run(vec![go(Step::Continue, &[10])], &[]));
        assert_eq!(p.reason, Reason::Breakpoint);
        let total = |p: &Pause| p.globals.iter().find(|v| v.name == "total").map(|v| v.value.clone()).unwrap_or_default();
        assert_eq!(total(&p), "0");
        let p = paused(run(vec![go(Step::Continue, &[10]), go(Step::Continue, &[10])], &[]));
        assert_eq!(total(&p), "2");
        let end = run(vec![go(Step::Continue, &[10]), go(Step::Continue, &[])], &[]);
        match end {
            Outcome::Finished(Ok(_)) => {}
            other => panic!("{other:?}"),
        }
    }

    #[test]
    fn watches_and_immediate_statements_see_the_paused_procedure() {
        let script = vec![
            go(Step::Continue, &[10]),
            Command::Immediate("i = 7".into()),
            Command::Immediate("?i * 2".into()),
            Command::Immediate("Debug.Print \"x\"".into()),
        ];
        let p = paused(run(script, &["i + 1", "nope("]));
        assert_eq!(p.immediate, ["", "14", "x"]);
        assert_eq!(p.watches[0], ("i + 1".to_string(), Ok("8".to_string())));
        assert!(p.watches[1].1.is_err());
        // The assignment stays in force when the run goes on.
        let script = vec![go(Step::Continue, &[10]), Command::Immediate("Total = 100".into()), go(Step::Continue, &[])];
        match run(script, &[]) {
            Outcome::Finished(Ok(_)) => {}
            other => panic!("{other:?}"),
        }
    }

    #[test]
    fn stop_pauses_only_under_the_debugger_and_ends_a_normal_run() {
        let p = Project::from_sources(vec![SourceModule {
            name: "M".into(),
            kind: ModuleKind::Standard,
            source: "Sub A()\n    Debug.Print 1\n    Stop\n    Debug.Print 2\nEnd Sub\n".into(),
        }]);
        let mut host = NoHost;
        let mut vm = Vm::new(&p, &mut host);
        let _ = vm.run("A", vec![]);
        assert_eq!(vm.debug, ["1"]);
        let mut vm = Vm::new(&p, &mut host);
        let Outcome::Paused(first) = vm.run_debug("A", vec![], Plan { script: vec![Command::Go(Step::Continue, BTreeSet::new())], watches: vec![] })
        else {
            panic!()
        };
        assert_eq!((first.reason, first.frames[0].line, first.output.len()), (Reason::Stop, 3, 1));
    }
}
