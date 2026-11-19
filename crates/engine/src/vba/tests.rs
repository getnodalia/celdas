//! The Excel object model behind VBA, driven through `vba.run` on workbooks that carry synthetic
//! projects. The workbook a person owns is checked in `a_real_workbook_…`, which runs only when
//! its path is in `CELDAS_TURBO_XLSM`.

use std::sync::Arc;

use celdas_vba::{ModuleKind, SourceModule, project::write_bin};
use serde_json::{Value as Json, json};

use crate::{EngineError, Session};

fn session(code: &str) -> Session {
    let mut s = Session::new();
    s.new_workbook();
    let bin = write_bin(&[SourceModule { name: "Module1".into(), kind: ModuleKind::Standard, source: format!("Option Explicit\n{code}") }]).unwrap();
    Arc::make_mut(&mut s.doc_mut().unwrap().wb).vba_project = Some(Arc::new(bin));
    s
}

fn run(s: &mut Session, name: &str) -> Json {
    s.execute("vba.run", json!({"procedure": name})).unwrap_or_else(|e| panic!("{name}: {e}"))
}

fn text(s: &Session, sheet: usize, a1: &str) -> String {
    let d = s.doc().unwrap();
    let sh = d.wb.sheet(sheet).unwrap();
    crate::display::value_text(&d.wb, sh, celdas_core::CellRef::parse(a1).unwrap())
}

fn temp_dir(tag: &str) -> std::path::PathBuf {
    let dir = std::env::temp_dir().join(format!("celdas-vba-{tag}-{}", std::process::id()));
    let _ = std::fs::remove_dir_all(&dir);
    std::fs::create_dir_all(&dir).unwrap();
    dir
}

#[test]
fn the_project_is_listed_and_its_source_shown() {
    let mut s = session("Sub Hola()\nEnd Sub\nFunction Doble(x)\nDoble = x * 2\nEnd Function\n");
    let l = s.execute("vba.list", json!({})).unwrap();
    assert_eq!(l["present"], true);
    let names: Vec<&str> = l["procedures"].as_array().unwrap().iter().map(|p| p["name"].as_str().unwrap()).collect();
    assert_eq!(names, ["Hola", "Doble"]);
    assert_eq!(l["procedures"][1]["kind"], "function");
    assert_eq!(l["procedures"][1]["takesArguments"], true);
    let src = s.execute("vba.source", json!({"module": "module1"})).unwrap();
    assert!(src["modules"][0]["source"].as_str().unwrap().contains("Sub Hola()"));
    assert!(s.execute("vba.source", json!({"module": "nope"})).is_err());
    let mut blank = Session::new();
    blank.new_workbook();
    let none = blank.execute("vba.list", json!({})).unwrap();
    assert_eq!(none["present"], false);
}

#[test]
fn cells_are_written_read_and_calculated() {
    let mut s = session(
        r#"
Sub Llena()
    Dim i As Long, ws As Worksheet
    Set ws = ActiveSheet
    For i = 1 To 5
        ws.Cells(i, 1).Value = i * 10
        ws.Cells(i, 2).Value = "n" & i
    Next i
    ws.Range("C1").Formula = "=SUM(A1:A5)"
    ws.Range("D1").Value = ws.Range("C1").Value + 1
    ws.Range("A7:B8").Value = "x"
    ws.Range("E1:G1").Value = Array(1, "dos", True)
End Sub
"#,
    );
    let r = run(&mut s, "Llena");
    assert!(r["steps"].as_u64().unwrap() > 10);
    assert_eq!(text(&s, 0, "A5"), "50");
    assert_eq!(text(&s, 0, "B3"), "n3");
    assert_eq!(text(&s, 0, "C1"), "150");
    assert_eq!(text(&s, 0, "D1"), "151");
    assert_eq!((text(&s, 0, "A7"), text(&s, 0, "B8")), ("x".into(), "x".into()));
    assert_eq!((text(&s, 0, "E1"), text(&s, 0, "F1"), text(&s, 0, "G1")), ("1".into(), "dos".into(), "TRUE".into()));
}

#[test]
fn a_cell_by_cell_loop_goes_in_as_blocks_and_one_undo_step() {
    let mut s =
        session("Sub Llena()\nDim i As Long, j As Long\nFor i = 1 To 200\nFor j = 1 To 20\nCells(i, j).Value = i + j\nNext j\nNext i\nEnd Sub\n");
    let before = s.doc().unwrap().undo.len();
    run(&mut s, "Llena");
    assert_eq!(text(&s, 0, "T200"), "220");
    assert_eq!(s.doc().unwrap().undo.len(), before + 1, "the run is a single undo step");
    s.execute("edit.undo", json!({})).unwrap();
    assert_eq!(text(&s, 0, "T200"), "");
}

#[test]
fn a_failing_macro_changes_nothing_and_says_where() {
    let mut s = session("Sub Mal()\nRange(\"A1\").Value = 5\nDim x As Long\nx = 1 / 0\nEnd Sub\n");
    let e = s.execute("vba.run", json!({"procedure": "Mal"})).unwrap_err().to_string();
    assert!(e.contains("Division by zero") && e.contains("line 5"), "{e}");
    assert_eq!(text(&s, 0, "A1"), "");
    assert!(s.execute("vba.run", json!({"procedure": "NoExiste"})).unwrap_err().to_string().contains("not defined"));
    assert!(s.execute("vba.run", json!({})).is_err());
}

#[test]
fn formatting_merges_and_sizes() {
    let mut s = session(
        r#"
Sub Da_Formato()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(1)
    ws.Range("A1:C1").Merge
    ws.Range("A1").Value = "Titulo"
    With ws.Range("A1:C1")
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlTop
        .Font.Bold = True
        .Font.Size = 14
        .Font.Color = RGB(255, 0, 0)
        .Interior.Color = RGB(226, 239, 218)
        .RowHeight = 22
    End With
    ws.Range("A2:C3").NumberFormat = "0.00"
    ws.Range("A2").Value = 3.14159
    With ws.Range("A2:C3").Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    ws.Range("A3:C3").Borders(xlEdgeBottom).Weight = xlThick
    ws.Columns("A:C").AutoFit
    ws.Range("E1").Value = "x"
    ws.Range("E1").ClearFormats
    ws.Range("A1:C1").UnMerge
End Sub
"#,
    );
    run(&mut s, "Da_Formato");
    let d = s.doc().unwrap();
    let sh = d.wb.sheet(0).unwrap();
    let st = |a1: &str| d.wb.styles.get(sh.style_id(celdas_core::CellRef::parse(a1).unwrap())).clone();
    let a1 = st("A1");
    assert_eq!(a1.align.h, celdas_model::HAlign::Center);
    assert_eq!(a1.align.v, celdas_model::VAlign::Top);
    assert!(a1.font.bold);
    assert_eq!(a1.font.size, 14.0);
    assert_eq!(a1.font.color, celdas_model::Color::rgb(255, 0, 0));
    assert_eq!(a1.fill, celdas_model::Fill::solid(celdas_model::Color::rgb(226, 239, 218)));
    assert!((sh.row_height(0) - 22.0 * 4.0 / 3.0).abs() < 0.01, "points become pixels: {}", sh.row_height(0));
    assert_eq!(st("A2").num_fmt.as_str(), "0.00");
    assert_eq!(text(&s, 0, "A2"), "3.14");
    let b = st("B2").borders;
    assert_eq!(b.top.style, celdas_model::BorderStyle::Thin);
    assert_eq!(b.left.style, celdas_model::BorderStyle::Thin);
    assert_eq!(st("B3").borders.bottom.style, celdas_model::BorderStyle::Thick);
    assert_eq!(st("E1"), celdas_model::Style::default());
    assert!(sh.merge_at(celdas_core::CellRef::parse("B1").unwrap()).is_none(), "UnMerge undid the merge");
}

#[test]
fn end_up_finds_the_last_row_and_used_range_counts() {
    let mut s = session(
        r#"
Function Ultima() As Long
    Dim ws As Worksheet
    Set ws = ActiveSheet
    ws.Range("A1").Value = 1
    ws.Range("A2").Value = 2
    ws.Range("A3").Value = 3
    ws.Range("A9").Value = 9
    Ultima = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
End Function
Sub Guarda()
    Range("C1").Value = Ultima()
    Range("C2").Value = Range("A1").End(xlDown).Row
    Range("C3").Value = Range("A1:B4").Rows.Count & "x" & Range("A1:B4").Columns.Count
    Range("C4").Value = Range("A1:B4").Count
    Range("C5").Value = Range("A1").Offset(2, 1).Address
    Range("C6").Value = Range("A1").Resize(2, 3).Address(False, False)
    Range("C7").Value = ActiveSheet.UsedRange.Address
End Sub
"#,
    );
    run(&mut s, "Guarda");
    assert_eq!(text(&s, 0, "C1"), "9");
    assert_eq!(text(&s, 0, "C2"), "3");
    assert_eq!(text(&s, 0, "C3"), "4x2");
    assert_eq!(text(&s, 0, "C4"), "8");
    assert_eq!(text(&s, 0, "C5"), "$B$3");
    assert_eq!(text(&s, 0, "C6"), "A1:C2");
    assert_eq!(text(&s, 0, "C7"), "$A$1:$C$9");
}

#[test]
fn sheets_are_found_ignoring_case_and_accents_and_can_be_added() {
    let mut s = session(
        r#"
Sub Hojas()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets.Add
    ws.Name = "Álabes"
    ThisWorkbook.Sheets("álabes").Range("A1").Value = "hola"
    Sheets(1).Activate
    Range("A1").Value = Sheets.Count
    Dim n As String
    For Each ws In ThisWorkbook.Worksheets
        n = n & ws.Name & ";"
    Next ws
    Range("A2").Value = n
End Sub
"#,
    );
    run(&mut s, "Hojas");
    let names: Vec<String> = s.doc().unwrap().wb.sheets.iter().map(|x| x.name.clone()).collect();
    assert_eq!(names.len(), 2);
    assert!(names.contains(&"Álabes".to_string()));
    let i = names.iter().position(|n| n == "Álabes").unwrap();
    assert_eq!(text(&s, i, "A1"), "hola");
    let first = 1 - i;
    assert_eq!(text(&s, first, "A1"), "2");
    assert_eq!(text(&s, first, "A2"), format!("{};{};", names[0], names[1]));
    let e = {
        let mut s2 = session("Sub X()\nSheets(\"nope\").Range(\"A1\").Value = 1\nEnd Sub\n");
        s2.execute("vba.run", json!({"procedure": "X"})).unwrap_err().to_string()
    };
    assert!(e.contains("Subscript out of range"), "{e}");
}

#[test]
fn a_clicked_shape_gives_its_name_and_text_to_the_macro() {
    let mut s = session(
        r#"
Sub Ir_A_Hoja()
    Dim shp As Shape, destino As String
    On Error Resume Next
    Set shp = ActiveSheet.Shapes(CStr(Application.Caller))
    destino = shp.TextFrame2.TextRange.Text
    On Error GoTo 0
    Sheets(Trim(destino)).Activate
End Sub
"#,
    );
    s.execute("home.insertSheet", json!({"name": "Destino"})).unwrap();
    s.execute("sheet.activate", json!({"sheet": "Sheet1"})).unwrap();
    s.execute("insert.shape", json!({"kind": "textBox", "at": "B2", "text": " Destino "})).unwrap();
    let id = s.doc().unwrap().wb.sheet(0).unwrap().shapes[0].id;
    crate::cmd::edit(&mut s, |cx| {
        let sh = cx.sheet_mut(0)?;
        sh.shapes[0].name = "Boton Destino".into();
        sh.shapes[0].macro_name = "Ir_A_Hoja".into();
        Ok(())
    })
    .unwrap();
    s.execute("vba.run", json!({"procedure": "Ir_A_Hoja", "caller": "Boton Destino"})).unwrap();
    let d = s.doc().unwrap();
    assert_eq!(d.wb.sheet(d.wb.active_sheet()).unwrap().name, "Destino");
    assert!(id > 0);
    // With no caller (run from the list) the shape lookup fails quietly and the sheet "" is not found.
    let e = s.execute("vba.run", json!({"procedure": "Ir_A_Hoja"})).unwrap_err().to_string();
    assert!(e.contains("Subscript out of range"), "{e}");
}

#[test]
fn a_macro_builds_a_chart() {
    let mut s = session(
        r#"
Sub Grafica()
    Dim ws As Worksheet, shp As Shape, cht As Chart, srs As Series, i As Long
    Set ws = ActiveSheet
    For i = 1 To 5
        ws.Cells(i, 1).Value = i
        ws.Cells(i, 2).Value = i * i
    Next i
    On Error Resume Next
    ws.Shapes("Mi_Grafico").Delete
    On Error GoTo 0
    Set shp = ws.Shapes.AddChart2(240, xlXYScatterSmoothNoMarkers, ws.Range("D2").Left, ws.Range("D2").Top, 300, 200)
    shp.Name = "Mi_Grafico"
    Set cht = shp.Chart
    Do While cht.SeriesCollection.Count > 0
        cht.SeriesCollection(1).Delete
    Loop
    Set srs = cht.SeriesCollection.NewSeries
    srs.Name = "Cuadrados"
    srs.XValues = ws.Range(ws.Cells(1, 1), ws.Cells(5, 1))
    srs.Values = ws.Range("B1:B5")
    cht.HasTitle = True
    cht.ChartTitle.Text = "Titulo " & srs.Name
    cht.HasLegend = False
End Sub
"#,
    );
    run(&mut s, "Grafica");
    run(&mut s, "Grafica");
    let d = s.doc().unwrap();
    let charts = &d.wb.sheet(0).unwrap().charts;
    assert_eq!(charts.len(), 1, "the second run replaced the chart");
    let c = &charts[0];
    assert_eq!(c.name, "Mi_Grafico");
    assert_eq!(c.kind, celdas_model::ChartKind::ScatterLines);
    assert_eq!(c.series.len(), 1);
    assert_eq!(c.series[0].values, "Sheet1!$B$1:$B$5");
    assert_eq!(c.series[0].categories.as_deref(), Some("Sheet1!$A$1:$A$5"));
    assert!(c.series[0].smooth);
    assert_eq!(c.title.as_deref(), Some("Titulo Cuadrados"));
    assert_eq!(c.legend, celdas_model::LegendPos::None);
    assert_eq!(c.anchor.cell, celdas_core::CellRef::parse("D2").unwrap());
    assert_eq!((c.anchor.width, c.anchor.height), (300.0, 200.0));
}

#[test]
fn files_are_written_whole_on_close_and_read_back() {
    let dir = temp_dir("files");
    let d = dir.to_string_lossy().replace('\\', "/");
    let mut s = session(&format!(
        r#"
Sub Archivos()
    Dim f As Integer, l As String, n As Long
    MkDir "{d}/sub"
    f = FreeFile
    Open "{d}/sub/datos.txt" For Output As #f
    Print #f, "uno"
    Print #f, "dos"
    Close #f
    Open "{d}/sub/datos.txt" For Append As #f
    Print #f, "tres"
    Close #f
    Open "{d}/sub/datos.txt" For Input As #f
    Do While Not EOF(f)
        Line Input #f, l
        n = n + 1
        Cells(n, 1).Value = l
    Loop
    Close #f
    Cells(1, 3).Value = Dir("{d}/sub/*.txt")
    Kill "{d}/sub/datos.txt"
    Cells(2, 3).Value = Dir("{d}/sub/*.txt")
End Sub
"#
    ));
    let r = run(&mut s, "Archivos");
    assert_eq!(r["files"].as_array().unwrap().len(), 2);
    assert_eq!((text(&s, 0, "A1"), text(&s, 0, "A2"), text(&s, 0, "A3")), ("uno".into(), "dos".into(), "tres".into()));
    assert_eq!(text(&s, 0, "C1"), "datos.txt");
    assert_eq!(text(&s, 0, "C2"), "");
    assert!(!dir.join("sub/datos.txt").exists());
    let _ = std::fs::remove_dir_all(&dir);
}

#[test]
fn a_failed_macro_leaves_no_half_written_file() {
    let dir = temp_dir("atomic");
    let d = dir.to_string_lossy().replace('\\', "/");
    let mut s = session(&format!("Sub Mal()\nOpen \"{d}/out.txt\" For Output As #1\nPrint #1, \"a\"\nError 5\nClose #1\nEnd Sub\n"));
    assert!(s.execute("vba.run", json!({"procedure": "Mal"})).is_err());
    assert_eq!(std::fs::read_to_string(dir.join("out.txt")).unwrap_or_default(), "");
    let _ = std::fs::remove_dir_all(&dir);
}

#[test]
fn a_new_workbook_is_filled_and_saved() {
    let dir = temp_dir("saveas");
    let d = dir.to_string_lossy().replace('\\', "/");
    let mut s = session(&format!(
        r#"
Sub Exporta()
    Dim wb As Workbook
    Set wb = Workbooks.Add
    wb.Sheets(1).Range("A1").Value = "Hola"
    wb.Sheets(1).Range("A2").Formula = "=1+2"
    wb.SaveAs "{d}/nuevo", xlOpenXMLWorkbook
    wb.Close
    Range("A1").Value = Workbooks.Count
End Sub
"#
    ));
    run(&mut s, "Exporta");
    assert_eq!(text(&s, 0, "A1"), "1", "the macro's own workbook is untouched except for what it wrote");
    let bytes = std::fs::read(dir.join("nuevo.xlsx")).unwrap();
    let (wb, _) = celdas_xlsx::read_xlsx(&bytes).unwrap();
    let sh = wb.sheet(0).unwrap();
    assert_eq!(sh.value(celdas_core::CellRef::parse("A1").unwrap()), celdas_core::Value::Text("Hola".into()));
    assert_eq!(sh.value(celdas_core::CellRef::parse("A2").unwrap()), celdas_core::Value::Number(3.0));
    assert_eq!(s.doc().unwrap().wb.sheets.len(), 1);
    let _ = std::fs::remove_dir_all(&dir);
}

#[test]
fn message_boxes_are_collected_and_application_settings_remembered() {
    let mut s = session(
        "Sub Avisa()\nApplication.ScreenUpdating = False\nApplication.DisplayAlerts = False\nIf Application.ScreenUpdating Then MsgBox \"mal\"\nMsgBox \"Listo \" & 3, vbInformation, \"Titulo\"\nApplication.ScreenUpdating = True\nEnd Sub\n",
    );
    let r = run(&mut s, "Avisa");
    assert_eq!(r["messages"], json!([{"title": "Titulo", "text": "Listo 3"}]));
}

#[test]
fn worksheet_functions_and_evaluate_use_the_calculation_engine() {
    let mut s = session(
        r#"
Sub Calcula()
    Range("A1").Value = 4
    Range("A2").Value = 6
    Range("B1").Value = Application.WorksheetFunction.Sum(Range("A1:A2"))
    Range("B2").Value = Application.WorksheetFunction.Pi()
    Range("B3").Value = Application.Evaluate("A1*A2")
    Range("B4").Value = Application.WorksheetFunction.Max(3, 9, 2)
End Sub
"#,
    );
    run(&mut s, "Calcula");
    assert_eq!(text(&s, 0, "B1"), "10");
    assert_eq!(&text(&s, 0, "B2")[..7], "3.14159");
    assert_eq!(text(&s, 0, "B3"), "24");
    assert_eq!(text(&s, 0, "B4"), "9");
}

#[test]
fn a_workbook_from_a_file_asks_before_its_macros_run() {
    let mut s = session("Sub Hola()\nRange(\"A1\").Value = 1\nEnd Sub\n");
    s.execute("file.saveAs", json!({"path": temp_dir("trust").join("libro.xlsm").to_string_lossy()})).unwrap();
    let saved = s.execute("file.save", json!({"base64": true})).ok();
    let _ = saved;
    let path = s.doc().unwrap().path.clone().unwrap();
    let mut t = Session::new();
    t.execute("file.open", json!({"path": path})).unwrap();
    let e = t.execute("vba.run", json!({"procedure": "Hola"})).unwrap_err();
    assert!(matches!(e, EngineError::Approve(_)), "{e}");
    t.execute("vba.run", json!({"procedure": "Hola", "approve": true})).unwrap();
    assert_eq!(text(&t, 0, "A1"), "1");
    // The approval holds for the session.
    t.execute("vba.run", json!({"procedure": "Hola"})).unwrap();
    let _ = std::fs::remove_dir_all(std::path::Path::new(&path).parent().unwrap());
}

/// The coursework workbook this was written for: its macros must reproduce what Excel left in it.
#[test]
fn a_real_workbook_reproduces_its_excel_results() {
    let Ok(path) = std::env::var("CELDAS_TURBO_XLSM") else { return };
    let bytes = std::fs::read(path).unwrap();
    let mut s = Session::new();
    s.execute("file.open", json!({"name": "turbo.xlsm", "base64": crate::io::base64_encode(&bytes)})).unwrap();
    let snapshot = |s: &Session| -> Vec<Vec<((u32, u32), celdas_core::Value)>> {
        s.doc()
            .unwrap()
            .wb
            .sheets
            .iter()
            .map(|sh| {
                sh.used_range()
                    .map(|u| u.iter().map(|c| (c, sh.value(c))).filter(|(_, v)| !v.is_empty()).map(|(c, v)| ((c.row, c.col), v)).collect())
                    .unwrap_or_default()
            })
            .collect()
    };
    let before = snapshot(&s);
    // The macros write their exports where the platform's Desktop is; the test removes what they wrote.
    let mut written: Vec<String> = Vec::new();
    for p in ["CalcularRebanadasEstator", "ProyectarPerfil_TXTyASC", "ColocarRebanadasSobreArco2D"] {
        let out = s.execute("vba.run", json!({"procedure": p, "approve": true})).unwrap_or_else(|e| panic!("{p}: {e}"));
        if let Some(files) = out.get("files").and_then(|f| f.as_array()) {
            written.extend(files.iter().filter_map(|f| f.as_str().map(str::to_string)));
        }
    }
    for f in &written {
        let _ = std::fs::remove_file(f);
    }
    let after = snapshot(&s);
    let mut numbers = 0;
    for (b, a) in before.iter().zip(&after) {
        let a: std::collections::HashMap<_, _> = a.iter().cloned().collect();
        for (k, v) in b {
            let now = a.get(k).unwrap_or_else(|| panic!("{k:?} was {v:?} and is now empty"));
            if let (celdas_core::Value::Number(p), celdas_core::Value::Number(q)) = (v, now) {
                assert!((p - q).abs() <= 1e-9 * p.abs().max(1.0), "{k:?}: {p} vs {q}");
                numbers += 1;
            }
        }
    }
    assert!(numbers > 1000, "{numbers} numbers compared");
    let named: usize = s.doc().unwrap().wb.sheets.iter().map(|sh| sh.charts.iter().filter(|c| c.name == "Grafico_Rebanadas_Arco").count()).sum();
    assert_eq!(named, 1, "the macro leaves exactly one of its chart");
}
