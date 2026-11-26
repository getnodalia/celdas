//! Running and debugging from the editor, and the Immediate window. Every call goes through the
//! engine's `vba.run`, `vba.debug` and `vba.immediate` commands by `macro_ui::call`, which also
//! turns their questions (approve this macro?) into a prompt whose buttons run the command again.
//! What a command answered waits in `MacroUi::vba_result`, and `Ide::consume` takes it in, so an
//! answer that came from a prompt's button lands in the editor the same way.

use serde_json::{Value as Json, json};

use super::{Bottom, Failure, Ide, ImmLine, Proc};
use crate::SheetApp;
use crate::macro_ui;

/// `(message without the site, (module, line))` for an engine error that ends with
/// ` (in Module.Proc, line N)`.
pub fn split_site(message: &str) -> (String, Option<(String, u32)>) {
    let Some((head, tail)) = message.rsplit_once(" (in ") else { return (message.to_string(), None) };
    let Some(inner) = tail.strip_suffix(')') else { return (message.to_string(), None) };
    let Some((place, line)) = inner.rsplit_once(", line ") else { return (message.to_string(), None) };
    let (Some((module, _)), Ok(line)) = (place.split_once('.'), line.trim().parse::<u32>()) else { return (message.to_string(), None) };
    (head.trim_end().to_string(), Some((module.to_string(), line)))
}

fn lines(v: &Json) -> Vec<String> {
    v.as_array().into_iter().flatten().filter_map(Json::as_str).map(str::to_string).collect()
}

impl Ide {
    /// The breakpoints as the engine takes them.
    pub fn breakpoint_list(&self) -> Json {
        Json::Array(self.breakpoints.iter().map(|(m, l)| json!({"module": m, "line": l})).collect())
    }

    /// What F5 runs: the Sub the caret is in, or the module's only Sub that can be started by hand.
    pub fn target(&self) -> Option<String> {
        if let Some(p) = self.procedure_at_caret().filter(|p| p.runnable()) {
            return Some(format!("{}.{}", self.current, p.name));
        }
        let hand: Vec<&Proc> = self.analysis.procs.iter().filter(|p| p.runnable() && p.public).collect();
        match hand.as_slice() {
            [only] => Some(format!("{}.{}", self.current, only.name)),
            _ => None,
        }
    }

    /// Saves, then runs `procedure` ("Module.Name") without stopping. While a macro is paused, F5
    /// continues it instead.
    pub fn run(&mut self, app: &mut SheetApp, procedure: &str) {
        if self.paused() {
            self.step(app, "continue");
            return;
        }
        if !self.save_all(app) {
            return;
        }
        self.failure = None;
        self.output.clear();
        self.status = None;
        macro_ui::call(app, "Run Macro", "vba.run", json!({"procedure": procedure}));
        self.consume(app);
    }

    /// Saves, then starts `procedure` paused before its first statement.
    pub fn debug(&mut self, app: &mut SheetApp, procedure: &str) {
        if self.paused() || !self.save_all(app) {
            return;
        }
        self.failure = None;
        self.output.clear();
        self.status = None;
        let params = json!({"action": "start", "procedure": procedure, "breakpoints": self.breakpoint_list(), "watches": self.watches});
        macro_ui::call(app, "Debug Macro", "vba.debug", params);
        self.consume(app);
    }

    /// Step Into, Step Over, Step Out or Continue (`kind`: into, over, out, continue).
    pub fn step(&mut self, app: &mut SheetApp, kind: &str) {
        if !self.paused() {
            return;
        }
        let params = json!({"action": "step", "kind": kind, "breakpoints": self.breakpoint_list(), "watches": self.watches});
        macro_ui::call(app, "Debug Macro", "vba.debug", params);
        self.consume(app);
    }

    /// Ends the debugged macro. `keep`: what it did so far stays (one undo step).
    pub fn stop(&mut self, app: &mut SheetApp, keep: bool) {
        if !self.paused() {
            return;
        }
        macro_ui::call(app, "Debug Macro", "vba.debug", json!({"action": "stop", "keep": keep}));
        self.consume(app);
    }

    /// Tells a paused macro about breakpoints or watches that changed.
    pub fn push_debug_inputs(&mut self, app: &mut SheetApp) {
        if !self.paused() {
            return;
        }
        let params = json!({"action": "state", "breakpoints": self.breakpoint_list(), "watches": self.watches});
        self.keep_view = true;
        macro_ui::call(app, "Debug Macro", "vba.debug", params);
        self.consume(app);
    }

    /// Sets or clears the breakpoint of `line` in the current module.
    pub fn toggle_breakpoint(&mut self, app: &mut SheetApp, line: u32) {
        let key = (self.current.clone(), line);
        if !self.breakpoints.remove(&key) {
            self.breakpoints.insert(key);
        }
        self.push_debug_inputs(app);
    }

    /// Runs the line in the Immediate box: in the paused procedure when a macro is paused, in the
    /// workbook as it is otherwise.
    pub fn submit_immediate(&mut self, app: &mut SheetApp) {
        let text = self.imm_input.trim().to_string();
        if text.is_empty() || self.imm_pending.is_some() {
            return;
        }
        self.imm_input.clear();
        self.imm_recall = None;
        self.imm_focus = true;
        if !self.paused() && !self.save_all(app) {
            self.imm_input = text;
            return;
        }
        self.imm_pending = Some(text.clone());
        if self.paused() {
            let params = json!({"action": "immediate", "text": text, "breakpoints": self.breakpoint_list(), "watches": self.watches});
            macro_ui::call(app, "Debug Macro", "vba.debug", params);
        } else {
            macro_ui::call(app, "Immediate", "vba.immediate", json!({"text": text}));
        }
        self.consume(app);
        // A question (approve) keeps the line waiting; the answer's result clears it.
    }

    /// Takes in what a VBA command answered since the last pass.
    pub fn consume(&mut self, app: &mut SheetApp) {
        let Some((id, v)) = app.macros.vba_result.take() else { return };
        if let Some(problem) = v["problem"].as_str() {
            if let Some(input) = self.imm_pending.take() {
                self.imm_log.push(ImmLine { input, output: vec![split_site(problem).0], error: true });
                self.bottom = Bottom::Immediate;
                self.bottom_open = true;
            } else {
                self.fail(problem);
            }
            return;
        }
        match id.as_str() {
            "vba.run" => {
                self.failure = None;
                self.output = lines(&v["output"]);
                if !self.output.is_empty() {
                    self.bottom = Bottom::Output;
                    self.bottom_open = true;
                }
                let steps = v["steps"].as_u64().unwrap_or(0).to_string();
                let name = v["procedure"].as_str().unwrap_or("").to_string();
                self.status = Some(app.ui.language.tr_fill("Ran “{name}” ({steps} steps)", &[("name", &name), ("steps", &steps)]));
            }
            "vba.immediate" => {
                let input = self.imm_pending.take().unwrap_or_default();
                self.imm_log.push(ImmLine { input, output: lines(&v["output"]), error: false });
                show_messages(app, &v["messages"]);
                self.bottom = Bottom::Immediate;
                self.bottom_open = true;
            }
            "vba.debug" => self.take_state(app, &v),
            _ => {}
        }
    }

    /// Takes in a debugger answer: where the macro paused, or how it ended.
    pub fn take_state(&mut self, app: &mut SheetApp, v: &Json) {
        let lang = app.ui.language;
        match v["state"].as_str() {
            Some("paused") => {
                let first = self.debug.is_none();
                let keep_view = self.keep_view;
                self.keep_view = false;
                let typed = self.imm_pending.take();
                self.debug = Some(v.clone());
                let frames = v["frames"].as_array().map_or(0, Vec::len);
                self.frame = if first { 0 } else { self.frame.min(frames.saturating_sub(1)) };
                self.output = lines(&v["output"]);
                self.status = None;
                if let Some(input) = &typed {
                    self.imm_log.push(ImmLine { input: input.clone(), output: lines(&v["immediate"]), error: false });
                    self.bottom = Bottom::Immediate;
                    self.bottom_open = true;
                } else {
                    if first {
                        self.bottom = Bottom::Locals;
                        self.bottom_open = true;
                    }
                    if !keep_view && let Some((module, line)) = self.paused_at() {
                        self.go_to(&module, line, 0);
                    }
                }
            }
            Some("finished") => {
                self.debug = None;
                self.failure = None;
                self.output = lines(&v["output"]);
                show_messages(app, &v["messages"]);
                if !self.output.is_empty() {
                    self.bottom = Bottom::Output;
                    self.bottom_open = true;
                }
                self.status = Some(lang.tr("The macro finished.").to_string());
            }
            Some("failed") => {
                self.debug = None;
                show_messages(app, &v["messages"]);
                let e = &v["error"];
                let message = e["message"].as_str().unwrap_or("").to_string();
                let text = match e["number"].as_i64().filter(|n| *n != 0) {
                    Some(n) => lang.tr_fill("Run-time error {number}: {message}", &[("number", &n.to_string()), ("message", &message)]),
                    None => message,
                };
                let at = e["at"]["module"].as_str().zip(e["at"]["line"].as_u64().and_then(|l| u32::try_from(l).ok()));
                self.set_failure(at.map(|(m, l)| (m.to_string(), l)), text);
            }
            Some("stopped") => {
                self.debug = None;
                self.status = Some(lang.tr("Debugging stopped.").to_string());
            }
            _ => {}
        }
    }

    /// A run that failed with `message` (the engine's text, with its site at the end).
    pub fn fail(&mut self, message: &str) {
        let (text, site) = split_site(message);
        self.set_failure(site, text);
    }

    fn set_failure(&mut self, site: Option<(String, u32)>, message: String) {
        let (module, line) = match site {
            Some((m, l)) => (self.module(&m).map_or(m, |x| x.name.clone()), l),
            None => (self.current.clone(), 0),
        };
        self.failure = Some(Failure { module: module.clone(), line, message });
        self.bottom = Bottom::Problems;
        self.bottom_open = true;
        if line > 0 {
            self.go_to(&module, line, 0);
        }
    }
}

/// What the macro's `MsgBox` calls said, as the message box of the window.
fn show_messages(app: &mut SheetApp, messages: &Json) {
    let list: Vec<&Json> = messages.as_array().map(|m| m.iter().collect()).unwrap_or_default();
    if list.is_empty() {
        return;
    }
    let title = list.iter().filter_map(|m| m["title"].as_str()).find(|t| !t.is_empty()).unwrap_or("Celdas");
    let text = list.iter().filter_map(|m| m["text"].as_str()).collect::<Vec<_>>().join("\n\n");
    app.message = Some((title.to_string(), text));
}

/// With the editor closed, an answer nobody takes in: a failure is a message.
pub fn flush(app: &mut SheetApp) {
    if app.ide.is_some() {
        return;
    }
    let Some((id, v)) = app.macros.vba_result.take() else { return };
    if let Some(problem) = v["problem"].as_str() {
        app.message = Some(("Visual Basic".into(), split_site(problem).0));
    } else if id == "vba.debug" && v["state"] == "failed" {
        let message = v["error"]["message"].as_str().unwrap_or("").to_string();
        app.message = Some(("Visual Basic".into(), message));
    }
}

#[cfg(test)]
mod split_tests {
    use super::split_site;

    #[test]
    fn the_site_is_taken_off_the_end_of_an_error() {
        assert_eq!(split_site("Division by zero (in Module1.Fill, line 12)"), ("Division by zero".to_string(), Some(("Module1".to_string(), 12))));
        assert_eq!(split_site("Plain message"), ("Plain message".to_string(), None));
        assert_eq!(split_site("Odd (in nowhere)"), ("Odd (in nowhere)".to_string(), None));
        assert_eq!(split_site("Odd (in M.P, line x)"), ("Odd (in M.P, line x)".to_string(), None));
    }
}
