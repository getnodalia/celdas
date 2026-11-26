//! The VBA editor (Automate › Visual Basic, Alt+F11, Macros › Edit on a VBA macro): a workspace
//! that takes the whole window below the title bar, like File does, and lays itself out as an
//! IDE: a project explorer with the module's outline, tabs, the code with a margin of line
//! numbers and breakpoints, squiggles under mistakes, suggestions as you type, hover help, find
//! and replace, and a panel below with Problems, the Immediate window, Locals, Watch, the call
//! stack and the program's output. Running a macro and debugging it (breakpoints, Step Into,
//! Step Over, Step Out, Continue, Stop) are buttons and the keys of the Visual Basic editor.
//!
//! Everything goes through the engine's `vba.*` commands (`vba.save`, `vba.run`, `vba.debug`,
//! `vba.analyze`, `vba.complete`, `vba.hover`, `vba.occurrences`, `vba.format`, `vba.immediate`),
//! so the editor does nothing an agent can't do. Each save is one undo step, and the file's
//! project (its `Attribute` lines, forms and everything Celdas does not read) is kept as it was.
//!
//! The editor works on drafts: the code a person typed and has not saved, by module name. A draft
//! equal to the saved code is dropped. Saving, running and debugging save first.

use std::collections::{BTreeMap, BTreeSet};
use std::hash::{Hash, Hasher};
use std::ops::Range;

use egui::Ui;
use serde_json::{Value as Json, json};

use crate::SheetApp;
use crate::macro_ui;

mod chrome;
mod editor;
pub mod model;
mod panels;
mod run;

pub use run::flush;

/// One module of the project, as the editor lists it.
#[derive(Clone, Debug, PartialEq)]
pub struct Module {
    pub name: String,
    pub class: bool,
    /// The code behind ThisWorkbook or a sheet: it can be edited but not removed.
    pub document: bool,
    /// The code a person edits (the module without its `Attribute` lines), lines ending in `\n`.
    pub code: String,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum Severity {
    Error,
    Warning,
    Hint,
}

/// A mistake the analysis found. Lines count from 1, columns in characters.
#[derive(Clone, Debug, PartialEq)]
pub struct Diag {
    pub line: u32,
    pub start: usize,
    pub end: usize,
    pub severity: Severity,
    pub message: String,
}

/// A procedure of the module being edited.
#[derive(Clone, Debug, PartialEq)]
pub struct Proc {
    pub name: String,
    pub kind: String,
    pub public: bool,
    pub line: u32,
    pub end_line: u32,
    pub col: usize,
    pub signature: String,
}

impl Proc {
    /// A Sub that can be started by hand: it takes no arguments.
    pub fn runnable(&self) -> bool {
        self.kind == "sub"
            && self.signature.split_once('(').and_then(|(_, rest)| rest.split_once(')')).is_some_and(|(args, _)| args.trim().is_empty())
    }
}

#[derive(Clone, Debug, Default, PartialEq)]
pub struct Analysis {
    pub diags: Vec<Diag>,
    pub procs: Vec<Proc>,
    pub option_explicit: bool,
}

/// One suggestion in the list that opens as you type.
#[derive(Clone, Debug, PartialEq)]
pub struct Item {
    pub label: String,
    pub kind: String,
    pub detail: String,
    pub doc: String,
    pub insert: String,
}

#[derive(Clone, Debug, PartialEq)]
pub struct Popup {
    /// Column of the word the suggestion replaces.
    pub from: usize,
    pub items: Vec<Item>,
    pub selected: usize,
    /// Where the list hangs (under the caret when it opened).
    pub at: egui::Pos2,
}

/// The call being written, with the parameter the caret is in.
#[derive(Clone, Debug, PartialEq)]
pub struct Signature {
    pub label: String,
    pub params: Vec<String>,
    pub active: usize,
    pub doc: String,
}

/// What hovering a name says.
#[derive(Clone, Debug, PartialEq)]
pub struct Card {
    pub title: String,
    pub doc: String,
    pub error: bool,
    pub at: egui::Pos2,
}

/// The panel below the code.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum Bottom {
    Problems,
    Immediate,
    Locals,
    Watch,
    Calls,
    Output,
}

/// Find and replace.
#[derive(Clone, Debug, Default)]
pub struct Find {
    pub open: bool,
    pub replace: bool,
    pub query: String,
    pub with: String,
    pub how: model::Match,
    /// The hit the caret is on (0-based).
    pub index: usize,
    /// Give the field the keyboard on the next frame.
    pub focus: bool,
}

/// A question that waits for an answer.
#[derive(Clone, Debug, PartialEq)]
pub enum Ask {
    Remove(String),
    Close,
}

/// A line of the Immediate window and what it printed.
#[derive(Clone, Debug, PartialEq)]
pub struct ImmLine {
    pub input: String,
    pub output: Vec<String>,
    pub error: bool,
}

/// A runtime error: where it stopped and what it said.
#[derive(Clone, Debug, PartialEq)]
pub struct Failure {
    pub module: String,
    pub line: u32,
    pub message: String,
}

/// The New Module form in the explorer.
#[derive(Clone, Debug, Default)]
pub struct Adding {
    pub name: String,
    pub class: bool,
}

/// The VBA editor's state (see the module docs). One per window.
pub struct Ide {
    /// The workbook (`DocState::uid`) it edits.
    pub doc: u64,
    /// The workbook has a VBA project.
    pub project: bool,
    pub mods: Vec<Module>,
    /// What the engine could not read: `(module, message)`.
    pub errors: Vec<(String, String)>,
    pub drafts: BTreeMap<String, String>,
    pub tabs: Vec<String>,
    pub current: String,
    /// The caret and the other end of the selection in the current module (char indices).
    pub selection: (usize, usize),
    /// The caret moves here (char index) on the next frame, and the view follows it.
    pub reveal: Option<usize>,
    pub focus_editor: bool,
    pub analysis: Analysis,
    analysed: Option<(String, u64)>,
    pub popup: Option<Popup>,
    pub signature: Option<Signature>,
    pub card: Option<Card>,
    card_key: Option<(String, u64, usize)>,
    /// Other places the name under the caret is written (byte ranges), for the highlight.
    pub occurrences: Vec<Range<usize>>,
    occ_key: Option<(String, u64, usize)>,
    pub breakpoints: BTreeSet<(String, u32)>,
    /// The debugger's state while a macro is paused (`vba.debug` answer).
    pub debug: Option<Json>,
    /// The frame whose locals the panel shows (0 is the innermost).
    pub frame: usize,
    pub failure: Option<Failure>,
    pub bottom: Bottom,
    pub bottom_open: bool,
    pub imm_input: String,
    pub imm_log: Vec<ImmLine>,
    pub imm_pending: Option<String>,
    pub imm_focus: bool,
    pub imm_recall: Option<usize>,
    pub watches: Vec<String>,
    pub watch_input: String,
    pub output: Vec<String>,
    pub find: Find,
    pub status: Option<String>,
    pub ask: Option<Ask>,
    pub adding: Option<Adding>,
    /// Outline of the explorer is folded.
    pub outline_folded: bool,
    /// A suggestion was asked for with Ctrl+Space.
    pub explicit: bool,
    /// The last selection the suggestions were worked out for.
    last_selection: (usize, usize),
    /// The next debugger answer is a refresh (a breakpoint or watch changed): the caret stays.
    keep_view: bool,
}

impl Ide {
    fn new(doc: u64) -> Ide {
        Ide {
            doc,
            project: false,
            mods: Vec::new(),
            errors: Vec::new(),
            drafts: BTreeMap::new(),
            tabs: Vec::new(),
            current: String::new(),
            selection: (0, 0),
            reveal: None,
            focus_editor: true,
            analysis: Analysis::default(),
            analysed: None,
            popup: None,
            signature: None,
            card: None,
            card_key: None,
            occurrences: Vec::new(),
            occ_key: None,
            breakpoints: BTreeSet::new(),
            debug: None,
            frame: 0,
            failure: None,
            bottom: Bottom::Problems,
            bottom_open: false,
            imm_input: String::new(),
            imm_log: Vec::new(),
            imm_pending: None,
            imm_focus: false,
            imm_recall: None,
            watches: Vec::new(),
            watch_input: String::new(),
            output: Vec::new(),
            find: Find::default(),
            status: None,
            ask: None,
            adding: None,
            outline_folded: false,
            explicit: false,
            last_selection: (usize::MAX, usize::MAX),
            keep_view: false,
        }
    }

    /// The module called `name` (any case).
    pub fn module(&self, name: &str) -> Option<&Module> {
        self.mods.iter().find(|m| m.name.eq_ignore_ascii_case(name))
    }

    /// The text of module `name` as the editor shows it: the draft, or the saved code.
    pub fn text_of(&self, name: &str) -> String {
        self.drafts.get(name).cloned().or_else(|| self.module(name).map(|m| m.code.clone())).unwrap_or_default()
    }

    /// Records `text` as module `name`'s text: a draft, unless it is what the project holds.
    pub fn set_text(&mut self, name: &str, text: String) {
        if self.module(name).is_some_and(|m| m.code == text) {
            self.drafts.remove(name);
        } else {
            self.drafts.insert(name.to_string(), text);
        }
    }

    /// The names of the modules whose draft differs from what the project holds.
    pub fn dirty(&self) -> Vec<String> {
        self.mods.iter().filter(|m| self.drafts.get(&m.name).is_some_and(|d| *d != m.code)).map(|m| m.name.clone()).collect()
    }

    pub fn has_unsaved_changes(&self) -> bool {
        !self.dirty().is_empty()
    }

    /// Whether a macro is paused in the debugger.
    pub fn paused(&self) -> bool {
        self.debug.is_some()
    }

    /// The module, line and procedure the paused macro is about to run, innermost frame.
    pub fn paused_at(&self) -> Option<(String, u32)> {
        let f = self.debug.as_ref()?["frames"].get(0)?;
        Some((f["module"].as_str()?.to_string(), u32::try_from(f["line"].as_u64()?).ok()?))
    }

    /// Opens (or switches to) module `name`'s tab.
    pub fn open_module(&mut self, name: &str) {
        let Some(m) = self.module(name) else { return };
        let name = m.name.clone();
        if !self.tabs.contains(&name) {
            self.tabs.push(name.clone());
        }
        if self.current != name {
            self.current = name;
            self.popup = None;
            self.signature = None;
            self.card = None;
            self.analysed = None;
            self.focus_editor = true;
        }
    }

    /// Closes a tab; the one beside it takes over.
    pub fn close_tab(&mut self, name: &str) {
        let Some(i) = self.tabs.iter().position(|t| t == name) else { return };
        self.tabs.remove(i);
        if self.current == name {
            let next = self.tabs.get(i.min(self.tabs.len().saturating_sub(1))).cloned();
            self.current.clear();
            if let Some(n) = next {
                self.open_module(&n);
            }
        }
    }

    /// Moves the caret to `line` of module `name` (and shows it).
    pub fn go_to(&mut self, name: &str, line: u32, col: usize) {
        self.open_module(name);
        let text = self.text_of(&self.current.clone());
        self.reveal = Some(model::char_at(&text, line, col));
        self.focus_editor = true;
    }

    /// Reads the project again (after a save, an add, a removal or a change from outside).
    pub fn reload(&mut self, app: &mut SheetApp) {
        let list = app.session.run("vba.list", json!({}));
        let source = app.session.run("vba.source", json!({}));
        self.errors.clear();
        match (list, source) {
            (Ok(list), Ok(source)) if list["present"] == true => {
                self.project = true;
                self.mods = source["modules"]
                    .as_array()
                    .into_iter()
                    .flatten()
                    .map(|m| Module {
                        name: m["name"].as_str().unwrap_or("").to_string(),
                        class: m["kind"] == "class",
                        document: m["document"] == true,
                        code: model::normalize(m["code"].as_str().unwrap_or("")),
                    })
                    .collect();
                for e in list["errors"].as_array().into_iter().flatten() {
                    self.errors.push((e["module"].as_str().unwrap_or("").to_string(), e["message"].as_str().unwrap_or("").to_string()));
                }
            }
            (Err(e), _) | (_, Err(e)) => {
                self.status = Some(crate::clean_error(&e));
            }
            (Ok(_), Ok(_)) => {
                self.project = false;
                self.mods.clear();
            }
        }
        // A draft of a module that is gone, or that equals the code now, is dropped.
        let mods = self.mods.clone();
        self.drafts.retain(|name, text| mods.iter().any(|m| m.name == *name && m.code != *text));
        self.tabs.retain(|t| mods.iter().any(|m| m.name == *t));
        if self.module(&self.current).is_none() {
            self.current.clear();
        }
        if self.current.is_empty() {
            let first = self.tabs.first().cloned().or_else(|| self.starting_module());
            if let Some(n) = first {
                self.open_module(&n);
            }
        }
        self.analysed = None;
    }

    /// A new empty Sub at the end of the first standard module (the project is created first when
    /// the workbook has none), with the caret inside it.
    pub fn new_macro(&mut self, app: &mut SheetApp) {
        if !self.project {
            match app.run("vba.create", json!({"code": chrome::STARTER})) {
                Ok(_) => {
                    self.reload(app);
                    self.go_to("Module1", 4, 0);
                }
                Err(e) => app.message = Some(("Visual Basic".into(), crate::clean_error(&e))),
            }
            return;
        }
        let Some(module) = self.starting_module() else { return };
        let taken: String = self.mods.iter().map(|m| self.text_of(&m.name).to_lowercase()).collect::<Vec<_>>().join("\n");
        let Some(name) = (1..1000).map(|i| format!("Macro{i}")).find(|n| !taken.contains(&format!("sub {}", n.to_lowercase()))) else { return };
        let mut text = self.text_of(&module);
        if !text.ends_with('\n') && !text.is_empty() {
            text.push('\n');
        }
        if !text.is_empty() {
            text.push('\n');
        }
        let line = model::line_count(&text) + 1;
        text.push_str(&format!("Sub {name}()\n    \nEnd Sub\n"));
        self.set_text(&module, text);
        self.go_to(&module, line, 4);
    }

    /// The module to show first: one with a Sub in it, else the first standard module.
    fn starting_module(&self) -> Option<String> {
        self.mods
            .iter()
            .find(|m| !m.document && !m.class && m.code.to_lowercase().contains("sub "))
            .or_else(|| self.mods.iter().find(|m| !m.document && !m.class))
            .or_else(|| self.mods.first())
            .map(|m| m.name.clone())
    }

    /// Saves every module with a draft. Returns whether all of them were saved.
    pub fn save_all(&mut self, app: &mut SheetApp) -> bool {
        let names = self.dirty();
        if names.is_empty() {
            return true;
        }
        let mut notes = Vec::new();
        let mut ok = true;
        for name in &names {
            let Some(code) = self.drafts.get(name).cloned() else { continue };
            match app.run("vba.save", json!({"module": name, "code": code})) {
                Ok(r) => {
                    // The saved code is what the editor shows now: the draft is the code.
                    if let Some(m) = self.mods.iter_mut().find(|m| m.name == *name) {
                        m.code = code;
                    }
                    self.drafts.remove(name);
                    for e in r["errors"].as_array().into_iter().flatten().filter_map(Json::as_str) {
                        notes.push(format!("{name}: {e}"));
                    }
                }
                Err(e) => {
                    ok = false;
                    notes.push(format!("{name}: {}", crate::clean_error(&e)));
                    break;
                }
            }
        }
        self.reload(app);
        self.failure = None;
        if notes.is_empty() {
            let n = names.len().to_string();
            self.status = Some(app.ui.language.tr_fill("Saved the code of {n} module(s). The rest of the project was kept as it was.", &[("n", &n)]));
        } else {
            // The code is saved even when Celdas can't read it; say what it could not.
            self.status = Some(notes.join("\n"));
        }
        ok
    }

    /// Works out the marks of the current module's text when it changed.
    pub fn analyse(&mut self, app: &mut SheetApp, text: &str) {
        let mut h = std::collections::hash_map::DefaultHasher::new();
        text.hash(&mut h);
        let key = (self.current.clone(), h.finish());
        if self.analysed.as_ref() == Some(&key) {
            return;
        }
        self.analysed = Some(key);
        if text.len() > 600_000 {
            self.analysis = Analysis::default();
            return;
        }
        let Ok(r) = app.session.run("vba.analyze", json!({"code": text, "module": self.current})) else { return };
        let lang = app.ui.language;
        let diags = r["diagnostics"]
            .as_array()
            .into_iter()
            .flatten()
            .map(|d| {
                let template = d["template"].as_str().unwrap_or("");
                let args: Vec<(String, String)> = d["args"]
                    .as_object()
                    .into_iter()
                    .flatten()
                    .map(|(k, v)| {
                        let text = v.as_str().unwrap_or("");
                        // The parser's own sentence travels as the `message` of a syntax error.
                        (k.clone(), if template == "Syntax error: {message}" && k == "message" { lang.tr_syntax(text) } else { text.to_string() })
                    })
                    .collect();
                let refs: Vec<(&str, &str)> = args.iter().map(|(k, v)| (k.as_str(), v.as_str())).collect();
                let message = if template.is_empty() { d["message"].as_str().unwrap_or("").to_string() } else { lang.tr_fill(template, &refs) };
                Diag {
                    line: u32::try_from(d["line"].as_u64().unwrap_or(1)).unwrap_or(1),
                    start: usize::try_from(d["start"].as_u64().unwrap_or(0)).unwrap_or(0),
                    end: usize::try_from(d["end"].as_u64().unwrap_or(0)).unwrap_or(0),
                    severity: match d["severity"].as_str() {
                        Some("error") => Severity::Error,
                        Some("warning") => Severity::Warning,
                        _ => Severity::Hint,
                    },
                    message,
                }
            })
            .collect();
        let procs = r["procedures"]
            .as_array()
            .into_iter()
            .flatten()
            .map(|p| Proc {
                name: p["name"].as_str().unwrap_or("").to_string(),
                kind: p["kind"].as_str().unwrap_or("").to_string(),
                public: p["public"] == true,
                line: u32::try_from(p["line"].as_u64().unwrap_or(1)).unwrap_or(1),
                end_line: u32::try_from(p["endLine"].as_u64().unwrap_or(1)).unwrap_or(1),
                col: usize::try_from(p["col"].as_u64().unwrap_or(0)).unwrap_or(0),
                signature: p["signature"].as_str().unwrap_or("").to_string(),
            })
            .collect();
        self.analysis = Analysis { diags, procs, option_explicit: r["optionExplicit"] == true };
    }

    /// The procedure the caret is in.
    pub fn procedure_at_caret(&self) -> Option<&Proc> {
        let text = self.text_of(&self.current);
        let (line, _) = model::line_col(&text, self.selection.0);
        self.analysis.procs.iter().find(|p| p.line <= line && line <= p.end_line.max(p.line))
    }

    /// Problems of every kind, for the panel and the status line: `(module, line, severity, text)`.
    pub fn problems(&self) -> Vec<(String, u32, Severity, String)> {
        let mut out = Vec::new();
        if let Some(f) = &self.failure {
            out.push((f.module.clone(), f.line, Severity::Error, f.message.clone()));
        }
        for d in &self.analysis.diags {
            out.push((self.current.clone(), d.line, d.severity, d.message.clone()));
        }
        for (module, message) in &self.errors {
            if !module.eq_ignore_ascii_case(&self.current) && !message.is_empty() {
                out.push((module.clone(), 1, Severity::Error, message.clone()));
            }
        }
        out
    }

    /// Everything the IDE keeps in step with the workbook: it closes when the workbook does,
    /// follows a debug session that was started or ended elsewhere, takes in results.
    fn sync(&mut self, app: &mut SheetApp) -> bool {
        if !app.session.documents().iter().any(|d| d.uid == self.doc) {
            return false;
        }
        let present = app.session.documents().iter().find(|d| d.uid == self.doc).is_some_and(|d| d.wb.vba_project.is_some());
        if present != self.project {
            self.reload(app);
        }
        // A session the engine ended (the workbook was closed under it) leaves nothing paused.
        if self.debug.is_some() && app.session.vba_debug.is_none() {
            self.debug = None;
        }
        self.consume(app);
        true
    }
}

/// Opens the editor: `{module}` selects a module, `{procedure}` ("Module.Name") its module and
/// puts the caret on it, `{line}` on a line of that module.
pub fn open(app: &mut SheetApp, params: &Json) {
    let Some(doc) = app.session.active().map(|d| d.uid) else { return };
    let mut ide = match app.ide.take() {
        Some(ide) if ide.doc == doc => ide,
        _ => Box::new(Ide::new(doc)),
    };
    ide.reload(app);
    let (module, procedure) = match params["procedure"].as_str().and_then(|p| p.split_once('.')) {
        Some((m, p)) => (Some(m.to_string()), Some(p.to_string())),
        None => (params["module"].as_str().map(str::to_string), params["procedure"].as_str().map(str::to_string)),
    };
    if let Some(m) = module {
        ide.open_module(&m);
    }
    if let Some(p) = procedure {
        // Put the caret on the procedure's name once its module has been read.
        let text = ide.text_of(&ide.current.clone());
        let scan = app.session.run("vba.analyze", json!({"code": text, "module": ide.current})).ok();
        let at = scan
            .as_ref()
            .and_then(|r| r["procedures"].as_array())
            .and_then(|ps| ps.iter().find(|x| x["name"].as_str().is_some_and(|n| n.eq_ignore_ascii_case(&p))));
        if let Some(x) = at {
            let line = u32::try_from(x["line"].as_u64().unwrap_or(1)).unwrap_or(1);
            let col = usize::try_from(x["col"].as_u64().unwrap_or(0)).unwrap_or(0);
            ide.go_to(&ide.current.clone(), line, col);
        }
    }
    if let Some(line) = params["line"].as_u64().and_then(|l| u32::try_from(l).ok()) {
        let m = ide.current.clone();
        ide.go_to(&m, line, 0);
    }
    if params["new"] == true {
        ide.new_macro(app);
    }
    // A macro paused while the editor was away: pick the session up again.
    if app.session.vba_debug.is_some() && ide.debug.is_none() {
        let params = json!({"action": "state"});
        if let Ok(v) = app.run("vba.debug", params) {
            ide.take_state(app, &v);
        }
    }
    // Step Into from the Macros dialog: start the procedure paused before its first statement.
    if params["debug"] == true
        && let Some(p) = params["procedure"].as_str()
    {
        ide.debug(app, p);
    }
    app.ide = Some(ide);
}

/// Whether the editor holds code that is not in the workbook yet.
pub fn has_unsaved_changes(app: &SheetApp) -> bool {
    app.ide.as_ref().is_some_and(|i| i.has_unsaved_changes())
}

/// Draws the editor in place of the ribbon and the sheet.
pub fn show(app: &mut SheetApp, ui: &mut Ui) {
    let Some(mut ide) = app.ide.take() else { return };
    if !ide.sync(app) {
        return;
    }
    let mut leave = false;
    chrome::keys(app, ui.ctx(), &mut ide, &mut leave);
    chrome::layout(app, ui, &mut ide, &mut leave);
    chrome::questions(app, ui.ctx(), &mut ide, &mut leave);
    if leave {
        // Leaving while a macro is paused leaves the workbook locked: the status bar offers Stop.
        return;
    }
    app.ide = Some(ide);
}

/// Alt+F11 or the window's close button asks first when there is unsaved code.
pub fn request_close(app: &mut SheetApp) {
    let Some(ide) = app.ide.as_mut() else { return };
    if ide.has_unsaved_changes() {
        ide.ask = Some(Ask::Close);
    } else {
        app.ide = None;
    }
}

/// The status bar while a VBA macro is paused and the editor is closed: what is paused, and Show
/// and Stop. Returns whether it drew.
pub fn status(app: &mut SheetApp, ui: &mut Ui) {
    let Some(name) = app.session.vba_debug.as_ref().filter(|d| !d.running()).map(|d| d.procedure.clone()) else { return };
    if app.ide.is_some() {
        return;
    }
    let t = crate::theme::Tokens::get(ui.ctx());
    ui.add_space(12.0);
    let (r, _) = ui.allocate_exact_size(egui::vec2(10.0, 10.0), egui::Sense::hover());
    ui.painter().circle_filled(r.center(), 4.5, t.breakpoint);
    ui.label(egui::RichText::new(crate::i18n::ui_fill(ui, "Debugging “{name}”", &[("name", &name)])).font(crate::theme::ui_font(12.5)).color(t.text));
    let lang = app.ui.language;
    let small = |ui: &mut Ui, label: &str| {
        ui.add(egui::Button::new(egui::RichText::new(lang.tr(label)).font(crate::theme::ui_font(12.0))).min_size(egui::vec2(40.0, 18.0)))
    };
    if small(ui, "Show").on_hover_text(crate::i18n::ui_tr(ui, "Open the macro in the editor")).clicked() {
        app.open_dialog("vbaEditor", json!({}));
    }
    if small(ui, "Stop").on_hover_text(crate::i18n::ui_tr(ui, "Stop debugging and keep the changes")).clicked() {
        macro_ui::call(app, "Debug", "vba.debug", json!({"action": "stop", "keep": true}));
    }
}

#[cfg(test)]
mod tests;
