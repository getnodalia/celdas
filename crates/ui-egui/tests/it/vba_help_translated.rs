//! What the Visual Basic editor explains is in Spanish too: the help of every name in the two
//! catalogs, and the sentences of the syntax errors.

use celdas_ui_egui::i18n::Language;

/// The quoted description at the end of a catalog line.
fn docs(catalog: &str) -> impl Iterator<Item = &str> {
    catalog.lines().filter(|l| !l.trim_start().starts_with('#')).filter_map(|l| {
        let l = l.trim_end().strip_suffix('"')?;
        l.rsplit_once('"').map(|(_, doc)| doc)
    })
}

#[test]
fn every_catalog_description_has_a_spanish_wording() {
    let catalogs = [include_str!("../../../engine/src/vba/excel.catalog"), include_str!("../../../vba/src/ide/vba.catalog")];
    let all: Vec<&str> = catalogs.iter().flat_map(|c| docs(c)).collect();
    assert!(all.len() > 400, "the catalogs' descriptions were not found ({})", all.len());
    let mut missing: Vec<&str> =
        all.into_iter().filter(|d| !d.is_empty() && !d.starts_with("VarType: ")).filter(|d| Language::Es.tr(d) == *d).collect();
    missing.sort_unstable();
    missing.dedup();
    assert!(missing.is_empty(), "no Spanish for: {missing:#?}");
}

#[test]
fn syntax_errors_read_in_spanish() {
    let es = Language::Es;
    assert_eq!(es.tr_syntax("an expression was expected"), "se esperaba una expresión");
    assert_eq!(es.tr_syntax("expected `Then`"), "se esperaba `Then`");
    assert_eq!(es.tr_syntax("`End Sub` is missing"), "falta `End Sub`");
    assert_eq!(es.tr_syntax("unexpected character `$`"), "carácter inesperado `$`");
    // What it does not know stays as the parser wrote it, and English is untouched.
    assert_eq!(es.tr_syntax("something new"), "something new");
    assert_eq!(Language::En.tr_syntax("an expression was expected"), "an expression was expected");
}
