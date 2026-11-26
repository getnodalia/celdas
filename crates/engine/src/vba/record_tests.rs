//! A command written as code does what the command does: each case runs the steps as commands in
//! one workbook, runs the recorded `Sub` in another, and compares every cell and its style.

use std::sync::Arc;

use celdas_vba::{ModuleKind, SourceModule, project::write_bin};
use serde_json::{Value as Json, json};

use super::record::{Statements, procedure_name, recorded_sub, statements, string_literal};
use crate::Session;

fn fresh(setup: &[(&str, Json)]) -> Session {
    let mut s = Session::new();
    s.new_workbook();
    for (id, p) in setup {
        s.execute(id, p.clone()).unwrap_or_else(|e| panic!("setup {id}: {e}"));
    }
    s
}

/// Every cell of the first sheets with its text, formula and style.
fn snapshot(s: &Session) -> Vec<String> {
    let d = s.doc().unwrap();
    let mut out = Vec::new();
    for (i, sh) in d.wb.sheets.iter().enumerate() {
        out.push(format!("sheet {i} {}", sh.name));
        for row in 0..14u32 {
            for col in 0..8u32 {
                let at = celdas_core::CellRef::new(row, col);
                let text = crate::display::value_text(&d.wb, sh, at);
                let style = format!("{:?}", d.wb.styles.get(sh.style_id(at)));
                out.push(format!("{} {text:?} {:?} {style}", at.a1(), sh.cell(at).and_then(|c| c.formula.clone())));
            }
        }
        out.push(format!("merges {:?}", sh.merges));
    }
    out
}

/// Runs `steps` as commands, and the same steps as recorded code, and says where they differ.
fn same(setup: &[(&str, Json)], steps: &[(&str, Json)]) {
    let mut by_command = fresh(setup);
    for (id, p) in steps {
        by_command.execute(id, p.clone()).unwrap_or_else(|e| panic!("{id}: {e}"));
    }
    let owned: Vec<(String, Json)> = steps.iter().map(|(i, p)| ((*i).to_string(), p.clone())).collect();
    let (sub, skipped) = recorded_sub("Rec", "", &owned);
    assert!(skipped.is_empty(), "skipped {skipped:?} in\n{sub}");
    let mut by_code = fresh(setup);
    let bin = write_bin(&[SourceModule { name: "Module1".into(), kind: ModuleKind::Standard, source: format!("Option Explicit\r\n{sub}") }]).unwrap();
    Arc::make_mut(&mut by_code.doc_mut().unwrap().wb).vba_project = Some(Arc::new(bin));
    by_code.execute("vba.run", json!({"procedure": "Rec"})).unwrap_or_else(|e| panic!("{e} in\n{sub}"));
    let (a, b) = (snapshot(&by_command), snapshot(&by_code));
    for (x, y) in a.iter().zip(&b) {
        assert_eq!(x, y, "the code differs from the commands:\n{sub}");
    }
    assert_eq!(a.len(), b.len(), "{sub}");
}

fn sel(range: &str) -> (&'static str, Json) {
    ("selection.set", json!({"range": range, "active": range.split([':', ',']).next().unwrap()}))
}

fn data() -> Vec<(&'static str, Json)> {
    vec![
        ("range.setValues", json!({"range": "A1", "values": [["Name", "Qty", "Price"], ["Pen", 3, 1.5], ["Ink", 12, 4.25], ["Pad", 7, 2]]})),
        ("cell.set", json!({"cell": "D2", "input": "=B2*C2"})),
    ]
}

#[test]
fn entering_values_and_formulas() {
    same(
        &[],
        &[
            ("cell.set", json!({"cell": "A1", "input": "hello \"world\""})),
            ("cell.set", json!({"cell": "A2", "input": 3.5})),
            ("cell.set", json!({"cell": "A3", "input": "=A2*2"})),
            ("cell.set", json!({"cell": "B1", "input": true})),
            ("cell.set", json!({"cell": "B2", "input": "caf\u{e9} \u{2013} \u{4e2d}"})),
            ("range.setValues", json!({"range": "C1", "values": [[1, 2], [3, "=C1+D1"]]})),
            sel("E5"),
            ("cell.set", json!({"input": "typed"})),
        ],
    );
}

#[test]
fn formatting_the_selection() {
    same(
        &data(),
        &[
            sel("A1:C1"),
            ("home.bold", json!({"on": true})),
            ("home.italic", json!({"on": true})),
            ("home.underline", json!({"on": true})),
            ("home.fontName", json!({"name": "Georgia"})),
            ("home.fontSize", json!({"size": 14})),
            ("home.fontColor", json!({"color": "#C00000"})),
            ("home.fillColor", json!({"color": "#FFFF00"})),
            ("home.alignCenter", json!({})),
            ("home.alignMiddle", json!({})),
            ("home.wrapText", json!({"on": true})),
            sel("C2:C4"),
            ("home.numberFormat", json!({"code": "0.00"})),
            sel("B2:B4"),
            ("home.numberFormat", json!({"format": "Currency"})),
            ("home.percent", json!({})),
            sel("A2"),
            ("home.bold", json!({"on": false})),
            ("home.fontColor", json!({"color": "auto"})),
            ("home.fillColor", json!({"color": "none"})),
        ],
    );
}

#[test]
fn borders_and_merges() {
    same(
        &data(),
        &[
            sel("A1:D4"),
            ("home.borders", json!({"preset": "all"})),
            sel("A1:D1"),
            ("home.borders", json!({"preset": "thickBottom"})),
            sel("A3:B4"),
            ("home.borders", json!({"preset": "outside", "style": "medium"})),
            sel("C3:D4"),
            ("home.borders", json!({"preset": "none"})),
            sel("A6:C6"),
            ("home.mergeCenter", json!({})),
        ],
    );
}

#[test]
fn rows_columns_and_cells() {
    same(
        &data(),
        &[
            ("home.insertRows", json!({"rows": "2:3"})),
            ("home.deleteColumns", json!({"cols": "B"})),
            sel("A2"),
            ("home.insertRows", json!({})),
            ("home.rowHeight", json!({"rows": "1:1", "points": 30})),
            ("home.columnWidth", json!({"cols": "A:A", "chars": 20})),
            ("home.autofitColumnWidth", json!({"cols": "B:C"})),
            ("home.hideRows", json!({"rows": "6:6"})),
            ("home.hideColumns", json!({"cols": "C:C"})),
            ("home.unhideColumns", json!({"cols": "C:C"})),
            sel("A8:B9"),
            ("home.insertCells", json!({"shift": "right"})),
            ("home.deleteCells", json!({"shift": "left"})),
        ],
    );
}

#[test]
fn clearing_and_filling() {
    same(
        &data(),
        &[
            sel("A2:C2"),
            ("home.bold", json!({"on": true})),
            ("edit.clearFormats", json!({})),
            sel("D2:D4"),
            ("edit.fillDown", json!({})),
            sel("A4:C4"),
            ("edit.clearContents", json!({})),
            sel("A1"),
            ("edit.clearAll", json!({})),
            ("range.fill", json!({"range": "F1:F3", "input": "x"})),
        ],
    );
}

#[test]
fn the_clipboard() {
    same(
        &data(),
        &[
            sel("A1:C4"),
            ("edit.copy", json!({})),
            sel("F1"),
            ("edit.paste", json!({})),
            sel("A2:A3"),
            ("edit.cut", json!({})),
            sel("G6"),
            ("edit.paste", json!({})),
            sel("B2:B4"),
            ("edit.copy", json!({})),
            sel("H1"),
            ("edit.pasteSpecial", json!({"what": "values"})),
            sel("A1"),
            ("edit.paste", json!({"at": "A9"})),
        ],
    );
}

#[test]
fn sorting_and_filtering() {
    same(&data(), &[sel("A1:D4"), ("data.sortAscending", json!({"column": "B"}))]);
    same(&data(), &[sel("A1:D4"), ("data.sortDescending", json!({"column": "C"}))]);
    same(
        &data(),
        &[sel("A1:D4"), ("data.sort", json!({"header": true, "keys": [{"column": "B", "order": "desc"}, {"column": "A", "order": "asc"}]}))],
    );
    same(&data(), &[sel("A1:D4"), ("data.filter", json!({}))]);
}

#[test]
fn sheets() {
    same(
        &data(),
        &[
            ("home.insertSheet", json!({"name": "Datos"})),
            ("cell.set", json!({"cell": "A1", "input": "on the new sheet"})),
            ("sheet.activate", json!({"sheet": "Sheet1"})),
            ("sheet.rename", json!({"name": "Ventas"})),
            ("sheet.activate", json!({"sheet": "Datos"})),
            ("home.deleteSheet", json!({})),
        ],
    );
}

#[test]
fn commands_with_no_statement_are_left_as_comments() {
    for (id, p) in [
        ("insert.chart", json!({})),
        ("home.increaseDecimal", json!({})),
        ("formulas.autoSum", json!({})),
        ("home.fontColor", json!({"color": {"theme": 4}})),
        ("edit.paste", json!({"text": "a\tb"})),
        ("data.sort", json!({"keys": [{"column": "A", "by": "cellColor"}]})),
    ] {
        assert!(matches!(statements(id, &p), Statements::Skipped), "{id}");
    }
    let (sub, skipped) = recorded_sub("Rec", "", &[("insert.chart".into(), json!({})), ("home.bold".into(), json!({"on": true}))]);
    assert_eq!(skipped, ["insert.chart"]);
    assert!(sub.contains("' Not recorded as code: insert.chart") && sub.contains("Selection.Font.Bold = True"));
}

#[test]
fn text_and_names_are_written_so_the_module_can_be_saved() {
    assert_eq!(string_literal("say \"hi\""), "\"say \"\"hi\"\"\"");
    assert_eq!(string_literal("a\nb"), "\"a\" & vbLf & \"b\"");
    assert_eq!(string_literal("\u{4e2d}"), "ChrW(20013)");
    assert_eq!(string_literal("\u{1f600}"), "ChrW(-10179) & ChrW(-8704)");
    assert_eq!(string_literal(""), "\"\"");
    assert_eq!(procedure_name("My macro-1"), "My_macro_1");
    assert_eq!(procedure_name("1st"), "M1st");
    // Hostile parameters give a comment, never a panic or a broken statement.
    for p in [json!({"range": "A1\"\nKill"}), json!({"range": 5}), json!({"sheet": {"a": 1}}), json!({"sheet": "x", "on": true})] {
        match statements("home.bold", &p) {
            Statements::Code(lines) => assert!(lines.iter().all(|l| !l.contains('\n')), "{lines:?}"),
            Statements::Skipped => {}
        }
    }
}

#[test]
fn recording_as_visual_basic_saves_a_sub_that_plays_back() {
    let mut s = fresh(&[]);
    let started = s.execute("macro.record", json!({"as": "vba", "name": "Formato", "description": "Headers in bold"})).unwrap();
    assert_eq!(started["as"], "vba");
    s.execute("cell.set", json!({"cell": "A1", "input": "Total"})).unwrap();
    s.execute("selection.set", json!({"range": "A1:C1", "active": "A1"})).unwrap();
    s.execute("home.bold", json!({"on": true})).unwrap();
    s.execute("home.fillColor", json!({"color": "#FFFF00"})).unwrap();
    s.execute("insert.chart", json!({})).ok();
    let done = s.execute("macro.record", json!({"stop": true})).unwrap();
    assert_eq!(done["saved"], "Formato", "{done}");
    assert_eq!(done["module"], "Module1");
    assert_eq!(done["as"], "vba");
    // The workbook has a VBA project now, with the Sub in it, and no macro of steps.
    let src = s.execute("vba.source", json!({"module": "Module1"})).unwrap();
    let code = src["modules"][0]["code"].as_str().unwrap();
    assert!(code.contains("Sub Formato()") && code.contains("' Headers in bold"), "{code}");
    assert!(code.contains("Range(\"A1\").Value = \"Total\"") && code.contains("Selection.Font.Bold = True"), "{code}");
    assert_eq!(s.execute("macro.list", json!({})).unwrap().as_array().unwrap().len(), 0);
    // It plays back: clear the sheet, run it, and the cells are as recorded.
    s.execute("selection.set", json!({"range": "A1:C1"})).unwrap();
    s.execute("edit.clearAll", json!({})).unwrap();
    s.execute("vba.run", json!({"procedure": "Formato"})).unwrap();
    let d = s.doc().unwrap();
    let sh = d.wb.sheet(0).unwrap();
    let a1 = celdas_core::CellRef::parse("A1").unwrap();
    assert_eq!(crate::display::value_text(&d.wb, sh, a1), "Total");
    assert!(d.wb.styles.get(sh.style_id(celdas_core::CellRef::parse("B1").unwrap())).font.bold);
    // A second recording with the same name goes in the next module, under a free name.
    s.execute("macro.record", json!({"as": "vba", "name": "Formato"})).unwrap();
    s.execute("home.italic", json!({"on": true})).unwrap();
    let again = s.execute("macro.record", json!({"stop": true})).unwrap();
    assert_eq!((again["saved"].as_str(), again["module"].as_str()), (Some("Formato_2"), Some("Module2")), "{again}");
    // Nothing recorded: an empty Sub, saved all the same.
    s.execute("macro.record", json!({"as": "vba"})).unwrap();
    let empty = s.execute("macro.record", json!({"stop": true})).unwrap();
    assert_eq!(empty["saved"], "Macro1");
    // `as` is checked.
    assert!(s.execute("macro.record", json!({"as": "python"})).is_err());
    s.execute("macro.record", json!({"cancel": true})).ok();
}
