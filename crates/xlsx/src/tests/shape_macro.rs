//! A shape keeps its name and the VBA procedure a click on it runs (`xdr:sp macro`).

use std::io::Read;

use celdas_model::Workbook;

use super::minimal;
use crate::{read_xlsx, write_xlsx};

const DRAWING: &str = r#"<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
<xdr:twoCellAnchor><xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>4</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>4</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
<xdr:sp macro="[0]!Ir_A_Hoja" textlink=""><xdr:nvSpPr><xdr:cNvPr id="7" name="Rectangle: Rounded Corners 6"/><xdr:cNvSpPr/></xdr:nvSpPr>
<xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1828800" cy="571500"/></a:xfrm><a:prstGeom prst="roundRect"><a:avLst/></a:prstGeom></xdr:spPr>
<xdr:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="es-ES"/><a:t>Diseño detallado</a:t></a:r></a:p></xdr:txBody></xdr:sp><xdr:clientData/></xdr:twoCellAnchor></xdr:wsDr>"#;

fn read_drawing(drawing: &str) -> Workbook {
    let rels = r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/></Relationships>"#;
    let bytes = minimal(
        r#"<sheetData/><drawing r:id="rId1"/>"#,
        &[("xl/worksheets/_rels/sheet1.xml.rels", rels), ("xl/drawings/drawing1.xml", drawing)],
        "",
        "",
    );
    read_xlsx(&bytes).unwrap().0
}

fn written_drawing(wb: &Workbook) -> String {
    let bytes = write_xlsx(wb).unwrap();
    let mut archive = zip::ZipArchive::new(std::io::Cursor::new(bytes)).unwrap();
    let mut drawing = String::new();
    archive.by_name("xl/drawings/drawing1.xml").unwrap().read_to_string(&mut drawing).unwrap();
    drawing
}

#[test]
fn a_shape_keeps_its_name_and_macro() {
    let wb = read_drawing(DRAWING);
    let shape = &wb.sheet(0).unwrap().shapes[0];
    assert_eq!(shape.name, "Rectangle: Rounded Corners 6");
    assert_eq!(shape.macro_name, "Ir_A_Hoja");

    let xml = written_drawing(&wb);
    assert!(xml.contains(r#"macro="[0]!Ir_A_Hoja""#), "{xml}");
    assert!(xml.contains(r#"name="Rectangle: Rounded Corners 6""#), "{xml}");
    let (again, _) = read_xlsx(&write_xlsx(&wb).unwrap()).unwrap();
    let shape = &again.sheet(0).unwrap().shapes[0];
    assert_eq!((shape.name.as_str(), shape.macro_name.as_str()), ("Rectangle: Rounded Corners 6", "Ir_A_Hoja"));
}

#[test]
fn macro_targets_lose_their_workbook_prefix() {
    for (attr, want) in [("[0]!Run", "Run"), ("'Book 1.xlsm'!Run", "Run"), ("Run", "Run"), ("", "")] {
        let wb = read_drawing(&DRAWING.replace("[0]!Ir_A_Hoja", attr));
        assert_eq!(wb.sheet(0).unwrap().shapes[0].macro_name, want, "{attr}");
    }
}
