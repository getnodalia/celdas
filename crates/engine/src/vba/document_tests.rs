//! The code behind ThisWorkbook and the sheets: a new project has it, it follows a renamed sheet,
//! `Me` and the sheet's code name reach the right sheet, and the names survive a save.

use serde_json::{Value as Json, json};

use crate::Session;

fn with_project(sheets: &[&str]) -> Session {
    let mut s = Session::new();
    s.new_workbook();
    for name in sheets {
        s.execute("home.insertSheet", json!({"name": name})).unwrap();
    }
    s.execute("vba.create", json!({"code": ""})).unwrap();
    s
}

fn text(s: &Session, sheet: usize, a1: &str) -> String {
    let d = s.doc().unwrap();
    let sh = d.wb.sheet(sheet).unwrap();
    crate::display::value_text(&d.wb, sh, celdas_core::CellRef::parse(a1).unwrap())
}

fn module<'a>(list: &'a Json, name: &str) -> &'a Json {
    list["modules"].as_array().unwrap().iter().find(|m| m["name"] == name).unwrap_or_else(|| panic!("no module {name}"))
}

#[test]
fn a_new_project_has_the_code_behind_the_workbook_and_every_sheet() {
    let mut s = with_project(&["Ventas"]);
    let list = s.execute("vba.list", json!({})).unwrap();
    for name in ["ThisWorkbook", "Sheet1", "Sheet2", "Module1"] {
        assert_eq!(module(&list, name)["document"], name != "Module1", "{name}");
    }
    let d = s.doc().unwrap();
    assert_eq!(d.wb.code_name.as_deref(), Some("ThisWorkbook"));
    let codes: Vec<_> = d.wb.sheets.iter().map(|sh| sh.code_name.clone().unwrap_or_default()).collect();
    assert_eq!(codes, ["Sheet1", "Sheet2"]);
    let e = s.execute("vba.remove", json!({"module": "Sheet1"})).unwrap_err();
    assert!(e.to_string().contains("can't be removed"), "{e}");
}

#[test]
fn a_macro_reaches_a_sheet_by_its_code_name_after_the_sheet_is_renamed() {
    let mut s = with_project(&["Ventas"]);
    s.execute("sheet.rename", json!({"sheet": "Sheet1", "name": "Resumen"})).unwrap();
    let code = "Sub Escribe()\r\n    Sheet2.Range(\"A1\").Value = 42\r\n    Sheet1.Range(\"A1\").Value = ThisWorkbook.Sheets.Count\r\nEnd Sub\r\n";
    s.execute("vba.save", json!({"module": "Module1", "code": code})).unwrap();
    s.execute("vba.run", json!({"procedure": "Escribe"})).unwrap();
    // The first sheet was renamed, the second is still Sheet2 in the code.
    assert_eq!(text(&s, 1, "A1"), "42");
    assert_eq!(text(&s, 0, "A1"), "2");
}

#[test]
fn me_is_the_sheet_the_code_lives_behind_not_the_active_one() {
    let mut s = with_project(&["Ventas"]);
    s.execute("vba.save", json!({"module": "Sheet2", "code": "Sub Rellena()\r\n    Me.Range(\"B2\").Value = \"aqui\"\r\nEnd Sub\r\n"})).unwrap();
    s.execute(
        "vba.save",
        json!({"module": "ThisWorkbook", "code": "Sub Cuenta()\r\n    Me.Sheets(1).Range(\"C3\").Value = Me.Sheets.Count\r\nEnd Sub\r\n"}),
    )
    .unwrap();
    // Sheet 1 is active while the code of sheet 2 runs.
    s.execute("vba.run", json!({"procedure": "Sheet2.Rellena"})).unwrap();
    s.execute("vba.run", json!({"procedure": "Cuenta"})).unwrap();
    assert_eq!(text(&s, 1, "B2"), "aqui");
    assert_eq!(text(&s, 0, "B2"), "");
    assert_eq!(text(&s, 0, "C3"), "2");
}

#[test]
fn a_copied_sheet_does_not_take_the_code_name_of_the_original() {
    let mut s = with_project(&[]);
    s.execute("sheet.move", json!({"sheet": "Sheet1", "copy": true, "to": 1})).unwrap();
    let d = s.doc().unwrap();
    assert_eq!(d.wb.sheets.first().and_then(|sh| sh.code_name.as_deref()), Some("Sheet1"));
    assert_eq!(d.wb.sheets.get(1).and_then(|sh| sh.code_name.as_deref()), None);
}

#[test]
fn code_names_and_the_project_survive_a_save() {
    let mut s = with_project(&["Ventas"]);
    s.execute("vba.save", json!({"module": "Sheet2", "code": "Sub Hola()\r\nEnd Sub\r\n"})).unwrap();
    let bytes = celdas_xlsx::write_xlsx_as(&s.doc().unwrap().wb, celdas_xlsx::XlsxKind::MacroWorkbook).unwrap();
    let (wb, _) = celdas_xlsx::read_xlsx(&bytes).unwrap();
    assert_eq!(wb.code_name.as_deref(), Some("ThisWorkbook"));
    let codes: Vec<_> = wb.sheets.iter().map(|sh| sh.code_name.clone()).collect();
    assert_eq!(codes, [Some("Sheet1".to_string()), Some("Sheet2".to_string())]);
    let project = celdas_vba::Project::from_bin(wb.vba_project.as_deref().unwrap()).unwrap();
    let names: Vec<_> = project.modules.iter().map(|m| m.name.as_str()).collect();
    assert_eq!(names, ["ThisWorkbook", "Sheet1", "Sheet2", "Module1"]);
    assert!(project.modules[2].source.contains("Sub Hola()"));
}
