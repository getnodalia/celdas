// SPDX-License-Identifier: MIT OR Apache-2.0
//! The editor's state without a window: drafts, new macros, running, debugging and breakpoints,
//! all through the engine's `vba.*` commands.

use celdas_engine::Session;
use serde_json::json;

use super::{Ide, open};
use crate::SheetApp;

const CODE: &str =
    "Option Explicit\r\n\r\nSub Fill()\r\n    Dim i As Long\r\n    For i = 1 To 3\r\n        Cells(i, 1).Value = i * 2\r\n    Next i\r\nEnd Sub\r\n";

fn app_with_project() -> SheetApp {
    let mut session = Session::default();
    session.new_workbook();
    let mut app = SheetApp::new(session, Default::default());
    app.run("vba.create", json!({"code": CODE})).expect("create");
    app
}

fn opened(app: &mut SheetApp, params: serde_json::Value) {
    open(app, &params);
    assert!(app.ide.is_some(), "the editor opens");
}

#[test]
fn opening_on_a_procedure_shows_its_module() {
    let mut app = app_with_project();
    opened(&mut app, json!({"procedure": "Module1.Fill"}));
    let ide = app.ide.as_ref().unwrap();
    assert!(ide.project);
    assert_eq!(ide.current, "Module1");
    assert_eq!(ide.tabs, ["Module1"]);
    assert!(ide.reveal.is_some(), "the caret moves to the procedure");
}

#[test]
fn a_workbook_without_a_project_opens_the_welcome() {
    let mut session = Session::default();
    session.new_workbook();
    let mut app = SheetApp::new(session, Default::default());
    opened(&mut app, json!({}));
    assert!(!app.ide.as_ref().unwrap().project);
}

#[test]
fn an_edit_is_a_draft_until_it_is_saved() {
    let mut app = app_with_project();
    opened(&mut app, json!({}));
    let mut ide = app.ide.take().unwrap();
    assert!(!ide.has_unsaved_changes());
    let text = format!("{}' note\n", ide.text_of("Module1"));
    ide.set_text("Module1", text.clone());
    assert_eq!(ide.dirty(), ["Module1"]);
    // Typing the original text back is not a change.
    let original = ide.module("Module1").unwrap().code.clone();
    ide.set_text("Module1", original);
    assert!(!ide.has_unsaved_changes());
    ide.set_text("Module1", text.clone());
    assert!(ide.save_all(&mut app));
    assert!(!ide.has_unsaved_changes());
    let source = app.session.run("vba.source", json!({})).unwrap();
    let module = source["modules"].as_array().and_then(|all| all.iter().find(|m| m["name"] == "Module1")).unwrap();
    assert!(module["code"].as_str().unwrap().contains("' note"));
}

#[test]
fn new_macros_get_names_that_are_free_even_before_saving() {
    let mut app = app_with_project();
    opened(&mut app, json!({"new": true}));
    let mut ide = app.ide.take().unwrap();
    assert!(ide.text_of("Module1").contains("Sub Macro1()"));
    ide.new_macro(&mut app);
    let text = ide.text_of("Module1");
    assert!(text.contains("Sub Macro2()"), "{text}");
    assert_eq!(text.matches("Sub Macro1()").count(), 1);
    assert!(ide.has_unsaved_changes());
}

#[test]
fn a_new_macro_creates_the_project_when_there_is_none() {
    let mut session = Session::default();
    session.new_workbook();
    let mut app = SheetApp::new(session, Default::default());
    opened(&mut app, json!({"new": true}));
    let ide = app.ide.as_ref().unwrap();
    assert!(ide.project);
    assert!(ide.text_of("Module1").contains("Sub MyMacro()"));
}

#[test]
fn running_from_the_editor_reports_the_run() {
    let mut app = app_with_project();
    opened(&mut app, json!({}));
    let mut ide = app.ide.take().unwrap();
    ide.run(&mut app, "Module1.Fill");
    assert!(ide.failure.is_none(), "{:?}", ide.failure.as_ref().map(|f| &f.message));
    let status = ide.status.clone().unwrap_or_default();
    assert!(status.contains("Fill"), "{status}");
    let a2 = app.session.run("cell.get", json!({"cell": "A2"})).ok();
    assert!(a2.is_some_and(|v| v.to_string().contains('4')));
}

#[test]
fn a_failing_run_marks_the_line() {
    let mut app = app_with_project();
    app.run("vba.save", json!({"module": "Module1", "code": "Sub Boom()\r\n    Dim x As Long\r\n    x = 1 / 0\r\nEnd Sub\r\n"})).expect("save");
    opened(&mut app, json!({}));
    let mut ide = app.ide.take().unwrap();
    ide.run(&mut app, "Module1.Boom");
    let failure = ide.failure.as_ref().expect("the run failed");
    assert_eq!((failure.module.as_str(), failure.line), ("Module1", 3));
    assert!(!failure.message.contains("(in "), "{}", failure.message);
}

#[test]
fn debugging_pauses_steps_and_stops() {
    let mut app = app_with_project();
    opened(&mut app, json!({"procedure": "Module1.Fill", "debug": true}));
    let mut ide = app.ide.take().unwrap();
    assert!(ide.paused(), "Step Into from the Macros dialog starts paused");
    let (module, first) = ide.paused_at().expect("paused somewhere");
    assert_eq!(module, "Module1");
    ide.step(&mut app, "over");
    let (_, second) = ide.paused_at().expect("still paused");
    assert_ne!(first, second, "a step moves the line");
    ide.stop(&mut app, false);
    assert!(!ide.paused());
}

#[test]
fn a_breakpoint_stops_a_debugged_run_and_shifts_with_edits() {
    let mut app = app_with_project();
    opened(&mut app, json!({}));
    let mut ide: Box<Ide> = app.ide.take().unwrap();
    ide.toggle_breakpoint(&mut app, 6);
    assert!(ide.breakpoints.contains(&("Module1".to_string(), 6)));
    ide.toggle_breakpoint(&mut app, 6);
    assert!(ide.breakpoints.is_empty());
    ide.toggle_breakpoint(&mut app, 6);
    ide.debug(&mut app, "Module1.Fill");
    ide.step(&mut app, "continue");
    assert_eq!(ide.paused_at(), Some(("Module1".to_string(), 6)));
    let locals = &ide.debug.as_ref().unwrap()["frames"][0]["locals"];
    assert!(locals.to_string().contains("\"i\""), "{locals}");
}

#[test]
fn the_immediate_window_runs_a_line() {
    let mut app = app_with_project();
    opened(&mut app, json!({}));
    let mut ide = app.ide.take().unwrap();
    ide.imm_input = "? 6 * 7".into();
    ide.submit_immediate(&mut app);
    let last = ide.imm_log.last().expect("an answer");
    assert_eq!(last.output.join(""), "42");
    assert!(!last.error);
}
