//! Stepping through a VBA macro: `vba.debug`.
//!
//! The interpreter cannot freeze a run between statements, so a pause ends the run and the next step
//! runs it again from the start, answering each earlier pause the way the person did
//! ([`celdas_vba::debug`]). This module is what makes that safe for a workbook: every replay begins
//! by rolling the workbook back to where the debug session started, so the run always sees the same
//! data and its writes are never applied twice. The session lives on the [`Session`] between calls
//! (one action per call) inside an undo group that spans them; stopping keeps the changes as one
//! undo step, or rolls them back. While a macro is paused its workbook is locked, as for
//! `macro.debug`: only the debugger, queries and the commands that look around run there.
//!
//! Lines are the lines of what an editor shows (a module's code without its `Attribute` header);
//! the header's lines are added and taken away at the edge.

use std::collections::{BTreeSet, HashMap};

use celdas_vba::debug::{Breakpoint, Command, Outcome, Pause, Plan, Reason, Step, Variable};
use celdas_vba::ide;
use celdas_vba::{Project, Vm};
use serde_json::{Value as Json, json};

use super::{ExcelHost, project_of, split_attributes, target};
use crate::cmd::str_param;
use crate::{EngineError, GroupToken, Result, Session};

/// The most breakpoints one call sets.
const MAX_BREAKPOINTS: usize = 10_000;
/// The most watch expressions.
const MAX_WATCHES: usize = 100;
/// The most answers one session keeps (each step replays them all).
const MAX_SCRIPT: usize = 20_000;
/// The longest text of one watch or Immediate line.
const MAX_TEXT: usize = 4_000;

const ACTIONS: [&str; 5] = ["start", "step", "immediate", "stop", "state"];

/// One answer to a pause, in editor lines.
#[derive(Clone)]
enum Act {
    Go(Step, BTreeSet<(String, u32)>),
    Immediate(String),
}

/// A VBA macro being debugged.
pub struct Debugger {
    pub doc: u64,
    pub procedure: String,
    token: Option<GroupToken>,
    /// Set while a replay runs (the run is out of the session's hands).
    running: bool,
    script: Vec<Act>,
    breakpoints: BTreeSet<(String, u32)>,
    watches: Vec<String>,
    /// The answer to the last call that paused, for `state`.
    last: Json,
}

impl Debugger {
    /// Whether edits of the workbook with this uid wait for the session to end.
    pub fn locks(&self, doc: u64) -> bool {
        self.doc == doc
    }

    /// Whether a replay is running (it edits the workbook itself).
    pub fn running(&self) -> bool {
        self.running
    }
}

/// Whether command `id` (a query when `query`) runs on a workbook this debugger locks.
pub fn runs_while_locked(id: &str, query: bool) -> bool {
    query || id == "vba.debug" || crate::macros::debug::LOOKS_AROUND.contains(&id)
}

/// Ends a session whose undo group is no longer open (something else closed it).
pub(crate) fn drop_if_ended(s: &mut Session) -> Option<String> {
    let d = s.vba_debug.as_ref()?;
    if d.running || d.token.as_ref().is_some_and(|t| s.group_is_open(t)) {
        return None;
    }
    s.vba_debug.take().map(|d| d.procedure)
}

/// After a panic inside a replay: the run lost its place, so it is rolled back.
pub(crate) fn recover(s: &mut Session) {
    if s.vba_debug.as_ref().is_some_and(Debugger::running)
        && let Some(mut d) = s.vba_debug.take()
        && let Some(token) = d.token.take()
    {
        let _ = s.group_rollback(token);
    }
}

fn bad(msg: impl Into<String>) -> EngineError {
    crate::cmd::bad("vba.debug", msg)
}

fn not_debugging() -> EngineError {
    EngineError::Other("No VBA macro is being debugged. Start one with action: start.".into())
}

fn already_ended(name: &str) -> String {
    format!("Macro '{name}' was no longer being debugged: its undo group had been closed outside the debugger, so nothing was kept or rolled back.")
}

/// `vba.debug {action: start|step|immediate|stop|state, procedure (start), approve?, caller?,
/// kind?: into|over|out|continue (step), text (immediate), keep? (stop), breakpoints?, watches?}`.
pub fn debug(s: &mut Session, p: &Json) -> Result<Json> {
    if s.depth() > 1 {
        return Err(EngineError::Disabled("vba.debug".into(), "The VBA debugger runs on its own, not as a step of a script or macro.".into()));
    }
    let ended = drop_if_ended(s);
    let action = match str_param(p, "action") {
        Some(a) => a.to_string(),
        None if s.vba_debug.is_some() || ended.is_some() => "state".to_string(),
        None => "start".to_string(),
    };
    if !ACTIONS.contains(&action.as_str()) {
        return Err(bad(format!("action must be start, step, immediate, stop or state, not '{action}'")));
    }
    if let Some(name) = ended
        && action != "start"
    {
        return Err(EngineError::Other(already_ended(&name)));
    }
    if action == "start" {
        return start(s, p);
    }
    let (doc, procedure) = s.vba_debug.as_ref().map(|d| (d.doc, d.procedure.clone())).ok_or_else(not_debugging)?;
    if s.active().map(|a| a.uid) != Some(doc) {
        return Err(EngineError::Other(format!("'{procedure}' is being debugged in another workbook; switch back to it first.")));
    }
    let breakpoints = p.get("breakpoints").filter(|v| !v.is_null()).map(breakpoints_param).transpose()?;
    let watches = p.get("watches").filter(|v| !v.is_null()).map(watches_param).transpose()?;
    let watches_changed = watches.is_some();
    if let Some(d) = s.vba_debug.as_mut() {
        if let Some(b) = breakpoints {
            d.breakpoints = b;
        }
        if let Some(w) = watches {
            d.watches = w;
        }
    }
    match action.as_str() {
        "step" => {
            let step = match str_param(p, "kind").unwrap_or("into") {
                "into" => Step::Into,
                "over" => Step::Over,
                "out" => Step::Out,
                "continue" => Step::Continue,
                other => return Err(bad(format!("`kind` must be into, over, out or continue, not '{other}'"))),
            };
            push(s, |d| Act::Go(step, d.breakpoints.clone()))?;
            replay(s)
        }
        "immediate" => {
            let text = str_param(p, "text").map(str::trim).filter(|t| !t.is_empty()).ok_or_else(|| bad("missing `text`"))?;
            if text.chars().count() > MAX_TEXT {
                return Err(bad("`text` is too long"));
            }
            let text = text.to_string();
            push(s, |_| Act::Immediate(text.clone()))?;
            replay(s)
        }
        "stop" => stop(s, p.get("keep").and_then(Json::as_bool).unwrap_or(true)),
        _ if watches_changed => replay(s),
        _ => {
            let mut out = s.vba_debug.as_ref().map(|d| d.last.clone()).ok_or_else(not_debugging)?;
            out["breakpoints"] = listed_breakpoints(s);
            Ok(out)
        }
    }
}

fn push(s: &mut Session, make: impl FnOnce(&Debugger) -> Act) -> Result<()> {
    let d = s.vba_debug.as_mut().ok_or_else(not_debugging)?;
    if d.script.len() >= MAX_SCRIPT {
        return Err(bad("this debug session has run too long; stop it and start again"));
    }
    let act = make(d);
    d.script.push(act);
    Ok(())
}

/// `[{module, line}]` (lines of the editor's code, 1-based).
fn breakpoints_param(v: &Json) -> Result<BTreeSet<(String, u32)>> {
    let list = v.as_array().ok_or_else(|| bad("`breakpoints` must be a list of {module, line}"))?;
    if list.len() > MAX_BREAKPOINTS {
        return Err(bad(format!("`breakpoints` lists {} lines; at most {MAX_BREAKPOINTS} can be set", list.len())));
    }
    let mut out = BTreeSet::new();
    for b in list {
        let module = b.get("module").and_then(Json::as_str).filter(|m| !m.trim().is_empty());
        let line = b.get("line").and_then(Json::as_u64).filter(|l| *l >= 1).and_then(|l| u32::try_from(l).ok());
        let (Some(module), Some(line)) = (module, line) else { return Err(bad("a breakpoint is {module: \"Module1\", line: 12}")) };
        out.insert((module.to_lowercase(), line));
    }
    Ok(out)
}

fn watches_param(v: &Json) -> Result<Vec<String>> {
    let list = v.as_array().ok_or_else(|| bad("`watches` must be a list of expressions"))?;
    if list.len() > MAX_WATCHES {
        return Err(bad(format!("at most {MAX_WATCHES} watches can be set")));
    }
    list.iter()
        .map(|w| match w.as_str().map(str::trim) {
            Some(t) if !t.is_empty() && t.chars().count() <= MAX_TEXT => Ok(t.to_string()),
            _ => Err(bad("a watch is a non-empty expression")),
        })
        .collect()
}

fn listed_breakpoints(s: &Session) -> Json {
    let list: Vec<Json> = s.vba_debug.iter().flat_map(|d| d.breakpoints.iter()).map(|(m, l)| json!({"module": m, "line": l})).collect();
    Json::Array(list)
}

fn start(s: &mut Session, p: &Json) -> Result<Json> {
    if let Some(d) = &s.vba_debug {
        return Err(EngineError::Other(format!("VBA macro '{}' is being debugged. Stop it first.", d.procedure)));
    }
    if let Some(d) = &s.macro_debug {
        return Err(EngineError::Other(format!("Macro '{}' is being debugged. Stop it first.", d.name)));
    }
    let (project, name) = target(s, p, "vba.debug")?;
    let breakpoints = p.get("breakpoints").filter(|v| !v.is_null()).map(breakpoints_param).transpose()?.unwrap_or_default();
    let watches = p.get("watches").filter(|v| !v.is_null()).map(watches_param).transpose()?.unwrap_or_default();
    // A procedure that takes arguments can't be started by hand (the Macros dialog doesn't list it).
    let takes_args = project
        .procedures()
        .into_iter()
        .any(|x| (x.name.eq_ignore_ascii_case(&name) || format!("{}.{}", x.module, x.name).eq_ignore_ascii_case(&name)) && x.has_params);
    if takes_args {
        return Err(EngineError::Other(format!("'{name}' takes arguments; call it from a macro without them to debug it.")));
    }
    let doc = s.doc()?.uid;
    s.vba_debug = Some(Debugger { doc, procedure: name, token: None, running: false, script: Vec::new(), breakpoints, watches, last: Json::Null });
    replay(s)
}

/// Ends the session without keeping anything: its group is rolled back.
fn abandon(s: &mut Session, why: EngineError) -> EngineError {
    if let Some(mut d) = s.vba_debug.take()
        && let Some(token) = d.token.take()
    {
        let _ = s.group_rollback(token);
    }
    why
}

/// Runs the macro again from the start with every answer so far, on the workbook as it was when the
/// session began, and says where it paused or how it ended.
fn replay(s: &mut Session) -> Result<Json> {
    let (acts, procedure, watches, old) = {
        let d = s.vba_debug.as_mut().ok_or_else(not_debugging)?;
        d.running = true;
        (d.script.clone(), d.procedure.clone(), d.watches.clone(), d.token.take())
    };
    if let Some(token) = old
        && let Err(e) = s.group_rollback(token)
    {
        return Err(abandon(s, e));
    }
    let token = match s.group_begin(&format!("Run VBA: {procedure}")) {
        Ok(t) => t,
        Err(e) => return Err(abandon(s, e)),
    };
    if let Some(d) = s.vba_debug.as_mut() {
        d.token = Some(token);
    }
    let project = match project_of(s) {
        Ok(Some((p, _))) => p,
        Ok(None) => return Err(abandon(s, EngineError::Other("The workbook has no VBA project.".into()))),
        Err(e) => return Err(abandon(s, e)),
    };
    let offsets = header_lines(&project);
    let plan = Plan { script: acts.iter().map(|a| command(a, &offsets)).collect(), watches };
    let (outcome, site, messages, written) = {
        let mut host = ExcelHost::new(s);
        let (outcome, site, output) = {
            let mut vm = Vm::new(&project, &mut host);
            let outcome = vm.run_debug(&procedure, Vec::new(), plan);
            (outcome, vm.error_site().cloned(), vm.debug.clone())
        };
        // A paused run's open files are dropped unwritten: only a run that ends writes them.
        let landed = match &outcome {
            Outcome::Finished(Ok(_)) => host.finish(),
            Outcome::Paused(_) => host.flush(),
            Outcome::Finished(Err(_)) => Ok(()),
        };
        let landed = landed.map_err(|e| e.to_string());
        (outcome, (site, landed, output), host.messages.clone(), host.written.clone())
    };
    let (site, landed, output) = site;
    match (outcome, landed) {
        (Outcome::Paused(pause), Ok(())) => {
            let out = paused_json(s, &project, &offsets, &procedure, &pause, &messages);
            if let Some(d) = s.vba_debug.as_mut() {
                d.running = false;
                d.last = out.clone();
            }
            Ok(out)
        }
        (Outcome::Finished(Ok(_)), Ok(())) => end(
            s,
            true,
            json!({"state": "finished", "procedure": procedure, "messages": message_list(&messages), "files": written, "output": output}),
        ),
        (Outcome::Finished(Err(e)), _) => {
            let at = site.map(|x| {
                let line = x.line.saturating_sub(offset_of(&offsets, &x.module));
                json!({"module": x.module, "procedure": x.procedure, "line": line})
            });
            let out = json!({
                "state": "failed",
                "procedure": procedure,
                "error": {"number": e.number, "message": e.description, "at": at},
                "messages": message_list(&messages),
            });
            end(s, false, out)
        }
        (_, Err(message)) => end(s, false, json!({"state": "failed", "procedure": procedure, "error": {"message": message}})),
    }
}

/// Ends the session: its changes become one undo step, or are rolled back.
fn end(s: &mut Session, keep: bool, mut out: Json) -> Result<Json> {
    let Some(mut d) = s.vba_debug.take() else { return Err(not_debugging()) };
    let ended = match d.token.take() {
        Some(token) if keep => s.group_end(token).is_ok(),
        Some(token) => s.group_rollback(token).is_ok(),
        None => false,
    };
    if !ended {
        return Err(EngineError::Other(already_ended(&d.procedure)));
    }
    out["kept"] = json!(keep);
    Ok(out)
}

fn stop(s: &mut Session, keep: bool) -> Result<Json> {
    let procedure = s.vba_debug.as_ref().map(|d| d.procedure.clone()).ok_or_else(not_debugging)?;
    end(s, keep, json!({"state": "stopped", "procedure": procedure}))
}

fn message_list(messages: &[super::Message]) -> Json {
    Json::Array(messages.iter().map(|m| json!({"title": m.title, "text": m.text})).collect())
}

/// How many lines of the header each module's source starts with, by lower-case module name.
fn header_lines(project: &Project) -> HashMap<String, u32> {
    project
        .modules
        .iter()
        .map(|m| {
            let header = split_attributes(&m.source).0;
            (m.name.to_lowercase(), u32::try_from(header.matches('\n').count()).unwrap_or(0))
        })
        .collect()
}

/// The ` (in Module.Proc, line N)` ending of a run's error, with the line as the editor counts
/// it: without the `Attribute` lines a module's source starts with.
pub(super) fn site_text(project: &Project, x: &celdas_vba::ErrorSite) -> String {
    let line = x.line.saturating_sub(offset_of(&header_lines(project), &x.module));
    format!(" (in {}.{}, line {line})", x.module, x.procedure)
}

fn offset_of(offsets: &HashMap<String, u32>, module: &str) -> u32 {
    offsets.get(&module.to_lowercase()).copied().unwrap_or(0)
}

fn command(a: &Act, offsets: &HashMap<String, u32>) -> Command {
    match a {
        Act::Immediate(t) => Command::Immediate(t.clone()),
        Act::Go(step, bps) => {
            let set: BTreeSet<Breakpoint> = bps.iter().map(|(m, l)| (m.clone(), l.saturating_add(offset_of(offsets, m)))).collect();
            Command::Go(*step, set)
        }
    }
}

/// The names as the code spells them: the debugger sees them in lower case.
struct Spelling {
    locals: HashMap<(String, String), HashMap<String, String>>,
    globals: HashMap<String, String>,
}

fn spelling(project: &Project) -> Spelling {
    let mut out = Spelling { locals: HashMap::new(), globals: HashMap::new() };
    for m in &project.modules {
        let scan = ide::scan(&m.source);
        for v in &scan.vars {
            let key = v.name.to_lowercase();
            match v.proc.and_then(|i| scan.procs.get(i)) {
                Some(p) => {
                    out.locals.entry((m.name.to_lowercase(), p.name.to_lowercase())).or_default().insert(key, v.name.clone());
                }
                None => {
                    out.globals.insert(key, v.name.clone());
                }
            }
        }
    }
    out
}

fn variables(list: &[Variable], names: Option<&HashMap<String, String>>) -> Json {
    Json::Array(
        list.iter()
            .map(|v| {
                let name = names.and_then(|n| n.get(&v.name)).unwrap_or(&v.name);
                json!({"name": name, "type": v.ty, "value": v.value})
            })
            .collect(),
    )
}

fn paused_json(s: &Session, project: &Project, offsets: &HashMap<String, u32>, procedure: &str, p: &Pause, messages: &[super::Message]) -> Json {
    let spell = spelling(project);
    let frames: Vec<Json> = p
        .frames
        .iter()
        .map(|f| {
            let key = (f.module.to_lowercase(), f.procedure.to_lowercase());
            json!({
                "module": f.module,
                "procedure": f.procedure,
                "line": f.line.saturating_sub(offset_of(offsets, &f.module)),
                "locals": variables(&f.locals, spell.locals.get(&key)),
            })
        })
        .collect();
    let watches: Vec<Json> = p
        .watches
        .iter()
        .map(|(expr, r)| match r {
            Ok(v) => json!({"expression": expr, "value": v}),
            Err(e) => json!({"expression": expr, "error": e}),
        })
        .collect();
    let reason = match p.reason {
        Reason::Step => "step",
        Reason::Breakpoint => "breakpoint",
        Reason::Stop => "stop",
    };
    json!({
        "state": "paused",
        "procedure": procedure,
        "reason": reason,
        "frames": frames,
        "globals": variables(&p.globals, Some(&spell.globals)),
        "watches": watches,
        "output": p.output,
        "immediate": p.immediate,
        "messages": message_list(messages),
        "breakpoints": listed_breakpoints(s),
    })
}

#[cfg(test)]
mod tests {
    use std::sync::Arc;

    use celdas_vba::{ModuleKind, SourceModule, project::write_bin};
    use serde_json::{Value as Json, json};

    use crate::{EngineError, Session};

    const CODE: &str = "Dim Total As Long\nSub Fill()\n    Dim i As Long\n    For i = 1 To 3\n        Cells(i, 1).Value = i * 10\n        Total = Total + i\n    Next i\n    Debug.Print Total\nEnd Sub\nSub Boom()\n    Cells(1, 1).Value = 5\n    Cells(2, 1).Value = 1 / 0\nEnd Sub\n";

    fn session() -> Session {
        let mut s = Session::new();
        s.new_workbook();
        let source = format!("Attribute VB_Name = \"Module1\"\r\n{CODE}");
        let bin = write_bin(&[SourceModule { name: "Module1".into(), kind: ModuleKind::Standard, source }]).unwrap();
        Arc::make_mut(&mut s.doc_mut().unwrap().wb).vba_project = Some(Arc::new(bin));
        s
    }

    fn line_of(text: &str) -> u64 {
        CODE.lines().position(|l| l.contains(text)).map(|i| i as u64 + 1).unwrap_or(0)
    }

    fn dbg(s: &mut Session, p: Json) -> Json {
        s.execute("vba.debug", p).unwrap_or_else(|e| panic!("{e}"))
    }

    fn cell(s: &Session, a1: &str) -> String {
        let d = s.doc().unwrap();
        let sh = d.wb.sheet(0).unwrap();
        crate::display::value_text(&d.wb, sh, celdas_core::CellRef::parse(a1).unwrap())
    }

    #[test]
    fn a_macro_pauses_at_its_first_statement_in_the_lines_of_the_editor() {
        let mut s = session();
        let r = dbg(&mut s, json!({"action": "start", "procedure": "Fill"}));
        assert_eq!(r["state"], "paused");
        assert_eq!(r["frames"][0]["procedure"], "Fill");
        assert_eq!(r["frames"][0]["line"], line_of("For i"));
        let r = dbg(&mut s, json!({"action": "step", "kind": "into"}));
        assert_eq!(r["frames"][0]["line"], line_of("Cells(i, 1)"));
        let i = r["frames"][0]["locals"].as_array().unwrap().iter().find(|v| v["name"] == "i").unwrap();
        assert_eq!(i["value"], "1");
    }

    #[test]
    fn each_step_replays_without_writing_twice_and_the_end_is_one_undo_step() {
        let mut s = session();
        let before = s.doc().unwrap().undo.len();
        let bp = json!([{"module": "Module1", "line": line_of("Total = Total + i")}]);
        let r = dbg(&mut s, json!({"action": "start", "procedure": "Fill", "breakpoints": bp}));
        assert_eq!(r["reason"], "step");
        let r = dbg(&mut s, json!({"action": "step", "kind": "continue"}));
        assert_eq!((r["reason"].as_str(), cell(&s, "A1").as_str(), cell(&s, "A2").as_str()), (Some("breakpoint"), "10", ""));
        let r = dbg(&mut s, json!({"action": "step", "kind": "continue"}));
        assert_eq!(r["reason"], "breakpoint");
        assert_eq!((cell(&s, "A1"), cell(&s, "A2")), ("10".to_string(), "20".to_string()));
        let total = r["globals"].as_array().unwrap().iter().find(|v| v["name"] == "Total").unwrap();
        assert_eq!(total["value"], "1");
        dbg(&mut s, json!({"action": "step", "kind": "continue", "breakpoints": []}));
        assert!(s.vba_debug.is_none());
        assert_eq!(cell(&s, "A3"), "30");
        assert_eq!(s.doc().unwrap().undo.len(), before + 1);
        s.execute("edit.undo", json!({})).unwrap();
        assert_eq!(cell(&s, "A1"), "");
    }

    #[test]
    fn the_workbook_is_locked_while_paused_and_stop_can_roll_it_back() {
        let mut s = session();
        dbg(&mut s, json!({"action": "start", "procedure": "Fill", "breakpoints": [{"module": "Module1", "line": line_of("Total = Total + i")}]}));
        let r = dbg(&mut s, json!({"action": "step", "kind": "continue"}));
        assert_eq!(r["state"], "paused");
        assert_eq!(cell(&s, "A1"), "10");
        let r = s.execute("cell.set", json!({"cell": "C1", "input": "x"}));
        assert!(matches!(&r, Err(EngineError::Disabled(..))), "{r:?}");
        assert!(s.execute("vba.run", json!({"procedure": "Fill"})).is_err());
        assert!(s.execute("vba.list", json!({})).is_ok());
        let r = dbg(&mut s, json!({"action": "stop", "keep": false}));
        assert_eq!(r["state"], "stopped");
        assert_eq!(cell(&s, "A1"), "");
        assert!(s.execute("cell.set", json!({"cell": "C1", "input": "x"})).is_ok());
    }

    #[test]
    fn immediate_lines_and_watches_see_the_paused_procedure() {
        let mut s = session();
        let bp = json!([{"module": "Module1", "line": line_of("Total = Total + i")}]);
        dbg(&mut s, json!({"action": "start", "procedure": "Fill", "breakpoints": bp, "watches": ["i * 2", "nope("]}));
        let r = dbg(&mut s, json!({"action": "step", "kind": "continue"}));
        assert_eq!(r["watches"][0]["value"], "2");
        assert!(r["watches"][1]["error"].is_string());
        let r = dbg(&mut s, json!({"action": "immediate", "text": "?i + 41"}));
        assert_eq!(r["immediate"][0], "42");
        let r = dbg(&mut s, json!({"action": "immediate", "text": "Total = 100"}));
        assert_eq!(r["state"], "paused");
        let r = dbg(&mut s, json!({"action": "state", "watches": ["Total"]}));
        assert_eq!(r["watches"][0]["value"], "100");
    }

    #[test]
    fn a_failing_macro_is_rolled_back_and_says_where() {
        let mut s = session();
        let r = dbg(&mut s, json!({"action": "start", "procedure": "Boom"}));
        assert_eq!(r["state"], "paused");
        let r = dbg(&mut s, json!({"action": "step", "kind": "continue"}));
        assert_eq!(r["state"], "failed");
        assert_eq!(r["error"]["at"]["line"], line_of("1 / 0"));
        assert_eq!(cell(&s, "A1"), "");
        assert!(s.vba_debug.is_none() && !s.in_group(s.doc().unwrap().uid));
    }

    #[test]
    fn closing_the_workbook_ends_the_session_and_bad_calls_are_refused() {
        let mut s = session();
        assert!(s.execute("vba.debug", json!({"action": "step"})).is_err());
        assert!(s.execute("vba.debug", json!({"action": "start", "procedure": "Nope"})).is_err());
        assert!(s.execute("vba.debug", json!({"action": "dance"})).is_err());
        dbg(&mut s, json!({"action": "start", "procedure": "Fill"}));
        assert!(s.execute("vba.debug", json!({"action": "start", "procedure": "Fill"})).is_err());
        assert!(s.execute("vba.debug", json!({"action": "step", "breakpoints": [{"line": 0}]})).is_err());
        s.close_document(0);
        assert!(s.vba_debug.is_none());
    }
}
