#![allow(clippy::unwrap_used, clippy::expect_used, clippy::panic)]

use super::*;

#[test]
fn saving_vba_to_a_macro_free_excel_file_requires_explicit_discard() {
    let mut s = Session::new();
    s.new_workbook();
    s.execute("cell.set", json!({"cell": "A1", "input": "keep"})).unwrap();
    std::sync::Arc::make_mut(&mut s.doc_mut().unwrap().wb).vba_project = Some(b"synthetic VBA".to_vec().into());
    let dir = std::env::temp_dir().join(format!("celdas-vba-save-{}", std::process::id()));
    std::fs::create_dir_all(&dir).unwrap();
    let xlsx = dir.join("book.xlsx");
    let xlsm = dir.join("book.xlsm");
    std::fs::write(&xlsx, b"existing").unwrap();
    assert!(s.execute("file.saveAs", json!({"path": xlsx})).is_err());
    assert_eq!(std::fs::read(&xlsx).unwrap(), b"existing");
    assert!(s.doc().unwrap().is_dirty());
    assert!(s.execute("file.saveBytes", json!({"format": "xltx"})).is_err());
    s.execute("file.saveAs", json!({"path": xlsm})).unwrap();
    let reopened = celdas_xlsx::read_xlsx(&std::fs::read(&xlsm).unwrap()).unwrap().0;
    assert_eq!(reopened.vba_project.as_deref().map(Vec::as_slice), Some(&b"synthetic VBA"[..]));
    s.execute("file.saveAs", json!({"path": xlsx, "discardVba": true})).unwrap();
    assert!(celdas_xlsx::read_xlsx(&std::fs::read(&xlsx).unwrap()).unwrap().0.vba_project.is_none());
    std::fs::remove_dir_all(dir).unwrap();
}
