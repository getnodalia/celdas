use celdas_engine::Session;

use super::*;
use crate::Services;

fn app() -> SheetApp {
    let mut session = Session::new();
    session.new_workbook();
    SheetApp::new(session, Services::default())
}

fn host(app: &mut SheetApp) -> CeldasHost<'_> {
    CeldasHost { app, lang: Lang::English }
}

fn cell(app: &mut SheetApp, at: &str) -> Value {
    app.run("cell.get", json!({"cell": at})).unwrap_or_default()
}

/// The tools the model is offered: every one named once, each schema an object, and only the
/// ones that leave the workbook's cells alone run without being asked about.
#[test]
fn the_tools_are_named_once_and_only_readers_run_unasked() {
    let mut a = app();
    let tools = host(&mut a).tools();
    let mut names: Vec<_> = tools.iter().map(|t| t.name.as_str()).collect();
    names.sort_unstable();
    let before = names.len();
    names.dedup();
    assert_eq!(before, names.len(), "a tool is listed twice");
    for t in &tools {
        assert_eq!(t.schema["type"], "object", "{}", t.name);
        assert!(!t.description.is_empty(), "{}", t.name);
    }
    for edit in ["write_range", "set_cell", "format_range", "insert_chart", "create_table", "sort_range", "filter", "run_command"] {
        assert!(tools.iter().any(|t| t.name == edit && !t.read_only), "{edit} must wait for the person");
    }
    for read in ["read_range", "get_cell", "evaluate_formula", "list_functions", "list_commands", "select"] {
        assert!(tools.iter().any(|t| t.name == read && t.read_only), "{read} is a reader");
    }
}

/// The assistant writes through the commands and a turn undoes in the number of steps it made.
#[test]
fn writing_cells_changes_the_workbook_and_undo_takes_it_back() {
    let mut a = app();
    let mut h = host(&mut a);
    let start = h.undo_depth();
    h.call("write_range", &json!({"range": "A1", "values": [["Item", "Cost"], ["Rent", 1200], ["Food", 300]]})).unwrap();
    h.call("set_cell", &json!({"cell": "B4", "input": "=SUM(B2:B3)"})).unwrap();
    let depth = h.undo_depth() - start;
    assert!(depth >= 2, "each write is a step, got {depth}");
    let read = h.call("read_range", &json!({"range": "A1:B4"})).unwrap();
    assert_eq!(read["values"][1], json!(["Rent", 1200.0]));
    assert_eq!(read["values"][3][1], json!(1500.0), "the formula was calculated");
    h.undo(depth).unwrap();
    let empty = h.call("read_range", &json!({"range": "A1:B4"})).unwrap();
    assert!(empty["values"].as_array().is_none_or(|rows| rows.iter().all(|r| r.as_array().is_none_or(|c| c.iter().all(Value::is_null)))));
}

/// A formula is checked without touching the sheet, and the snapshot shows what is in it.
#[test]
fn a_formula_is_evaluated_without_a_cell_and_the_snapshot_shows_the_data() {
    let mut a = app();
    let mut h = host(&mut a);
    h.call("write_range", &json!({"range": "A1", "values": [[1], [2], [3]]})).unwrap();
    let r = h.call("evaluate_formula", &json!({"formula": "SUM(A1:A3)*2"})).unwrap();
    assert_eq!(r["formula"], "=SUM(A1:A3)*2");
    assert_eq!(r["result"], json!(12.0));
    assert!(h.call("get_cell", &json!({"cell": "A4"})).is_ok());
    let snapshot = h.context().unwrap();
    assert_eq!(snapshot["preview"]["usedRange"], "A1:A3");
    assert_eq!(snapshot["preview"]["values"][2][0], json!(3.0));
    assert!(snapshot.get("undo").is_none(), "undo labels are not part of the snapshot");
}

/// The preview of a big sheet is its top-left corner and says so.
#[test]
fn a_big_sheet_is_previewed_by_its_corner() {
    let mut a = app();
    let mut h = host(&mut a);
    let rows: Vec<Value> = (0..(PREVIEW_ROWS + 40)).map(|r| json!([r, "x".repeat(300)])).collect();
    h.call("write_range", &json!({"range": "A1", "values": rows})).unwrap();
    let snapshot = h.context().unwrap();
    let shown = snapshot["preview"]["values"].as_array().unwrap().len() as u32;
    assert_eq!(shown, PREVIEW_ROWS);
    assert!(snapshot["preview"]["note"].is_string());
    let text = snapshot["preview"]["values"][0][1].as_str().unwrap();
    assert!(text.chars().count() <= LONG_CELL + 1, "long text is shortened");
}

/// Formatting, tables, sorting and charts go through the same commands as the ribbon.
#[test]
fn format_table_sort_and_chart_work() {
    let mut a = app();
    let mut h = host(&mut a);
    h.call("write_range", &json!({"range": "A1", "values": [["Name", "Score"], ["b", 2], ["a", 3], ["c", 1]]})).unwrap();
    let applied = h.call("format_range", &json!({"range": "A1:B1", "bold": true, "fillColor": "#FFE699"})).unwrap();
    assert_eq!(applied["applied"], json!(["home.bold", "home.fillColor"]));
    h.call("sort_range", &json!({"range": "A1:B4", "keys": ["B desc"], "header": true})).unwrap();
    let sorted = h.call("read_range", &json!({"range": "A2:A4"})).unwrap();
    assert_eq!(sorted["values"], json!([["a"], ["b"], ["c"]]));
    h.call("insert_chart", &json!({"range": "A1:B4", "type": "column", "title": "Scores"})).unwrap();
    h.call("create_table", &json!({"range": "A1:B4"})).unwrap();
    let info = h.call("read_range", &json!({"range": "A1"})).unwrap();
    assert_eq!(info["values"][0][0], "Name");
}

/// Files, windows, the UI and running macros are not the assistant's to touch.
#[test]
fn forbidden_commands_are_refused_and_hidden() {
    let mut a = app();
    let mut h = host(&mut a);
    for command in ["file.save", "file.open", "ui.dialog", "window.close", "app.quit", "file.print", "macro.run"] {
        let err = h.call("run_command", &json!({"command": command})).unwrap_err();
        assert!(err.contains("not available"), "{command}: {err}");
    }
    let listed = h.call("list_commands", &json!({})).unwrap();
    let commands = listed["commands"].as_array().unwrap();
    assert!(!commands.is_empty());
    assert!(commands.iter().all(|c| !forbidden(c["id"].as_str().unwrap_or_default())));
    let found = h.call("list_commands", &json!({"query": "sort"})).unwrap();
    assert!(found["commands"].as_array().unwrap().iter().any(|c| c["id"] == "data.sort"));
}

/// run_command acts on the selection like the ribbon button.
#[test]
fn run_command_acts_on_the_selection() {
    let mut a = app();
    let mut h = host(&mut a);
    h.call("write_range", &json!({"range": "A1", "values": [["hello"]]})).unwrap();
    h.call("select", &json!({"range": "A1"})).unwrap();
    h.call("run_command", &json!({"command": "home.bold", "params": {"on": true}})).unwrap();
    assert_eq!(cell(&mut a, "A1")["style"]["font"]["bold"], json!(true));
}

/// Bad input comes back as a message the model can act on, never a panic.
#[test]
fn bad_input_is_an_error_the_model_can_read() {
    let mut a = app();
    let mut h = host(&mut a);
    assert!(h.call("write_range", &json!({"range": "A1"})).is_err());
    assert!(h.call("set_cell", &json!({"cell": "A1"})).is_err());
    assert!(h.call("format_range", &json!({"range": "A1"})).unwrap_err().contains("nothing to apply"));
    assert!(h.call("format_range", &json!({"range": "A1", "horizontalAlign": "diagonal"})).is_err());
    assert!(h.call("sort_range", &json!({"keys": []})).is_err());
    assert!(h.call("run_command", &json!({"command": "no.such"})).is_err());
    assert!(h.call("nothing", &json!({})).unwrap_err().contains("no tool"));
    assert!(h.call("evaluate_formula", &json!({"formula": "=SUM("})).is_err());
    assert!(h.undo(1).is_err(), "nothing to undo in a new workbook");
}

/// The approval card and the suggestions speak both languages.
#[test]
fn descriptions_and_suggestions_are_in_both_languages() {
    let mut a = app();
    let mut h = host(&mut a);
    let write = json!({"range": "B2", "values": [[1, 2], [3, 4]]});
    assert_eq!(h.describe("write_range", &write).unwrap(), "Write 2 × 2 cells from B2");
    assert_eq!(h.describe("set_cell", &json!({"cell": "C3", "input": "=A1"})).unwrap(), "Enter “=A1” in C3");
    h.lang = Lang::Spanish;
    assert_eq!(h.describe("write_range", &write).unwrap(), "Escribir 2 × 2 celdas desde B2");
    assert!(h.describe("run_command", &json!({"command": "data.sort"})).unwrap().starts_with("Ejecutar "));
    assert!(h.describe("read_range", &json!({})).is_none());
    assert_eq!(h.suggestions(Lang::Spanish).len(), 4);
    assert_eq!(h.suggestions(Lang::English).len(), 4);
}

/// The panel and the button draw inside the real window, open and closed.
#[test]
fn the_button_and_the_panel_draw_in_a_frame() {
    let ctx = egui::Context::default();
    let mut a = app();
    let frame = |a: &mut SheetApp| {
        ctx.run_ui(egui::RawInput::default(), |ui| {
            a.logic(ui.ctx());
            a.ui(ui);
        })
        .drop_without_applying_deltas();
    };
    for _ in 0..4 {
        frame(&mut a);
    }
    assert!(a.assistant.is_some(), "the title bar made the assistant");
    for open in [true, false] {
        if let Some(x) = a.assistant.as_mut() {
            x.open = open;
        }
        frame(&mut a);
        frame(&mut a);
    }
}
