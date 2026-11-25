// SPDX-License-Identifier: MIT OR Apache-2.0
//! Every screen draws its text in the interface language. Each surface (a ribbon tab, a File page,
//! a dialog) is drawn in English and in Spanish; a label that reads the same in both and has a
//! Spanish catalog row was drawn without going through the catalog.

use std::collections::BTreeSet;

use celdas_ui_egui::SheetApp;
use celdas_ui_egui::backstage::Page;
use celdas_ui_egui::i18n::Language;
use celdas_ui_egui::ribbon::TABS;
use egui_kittest::kittest::NodeT as _;
use serde_json::json;

/// Everything the accessibility tree names on the current frame.
fn shown(h: &crate::common::Harness) -> BTreeSet<String> {
    let mut out = BTreeSet::new();
    for node in h.root().children_recursive() {
        let n = node.accesskit_node();
        for text in [n.label(), n.value()].into_iter().flatten() {
            let text = text.trim();
            if !text.is_empty() {
                out.insert(text.to_owned());
            }
        }
    }
    out
}

fn surface(language: Language, setup: &dyn Fn(&mut SheetApp)) -> BTreeSet<String> {
    let mut app = crate::common::app(crate::support::blank_session());
    app.ui.language = language;
    setup(&mut app);
    // Only the text is read here; an unnamed field is the accessibility tests' business.
    let mut h = crate::common::harness_with(crate::common::builder().with_size(egui::vec2(1200.0, 800.0)).with_accessibility_check(false), app, 6);
    h.run_steps(3);
    shown(&h)
}

/// Labels that read the same in English and Spanish on this surface.
fn same_in_both(setup: &dyn Fn(&mut SheetApp)) -> Vec<String> {
    let en = surface(Language::En, setup);
    let es = surface(Language::Es, setup);
    en.intersection(&es).cloned().collect()
}

/// Puts the app into the state of one surface.
type Setup = Box<dyn Fn(&mut SheetApp)>;

fn surfaces() -> Vec<(String, Setup)> {
    let mut v: Vec<(String, Setup)> = Vec::new();
    for tab in TABS {
        let tab = (*tab).to_owned();
        v.push((format!("tab {tab}"), Box::new(move |app: &mut SheetApp| app.ui.ribbon_tab = tab.clone())));
    }
    for page in Page::ALL {
        v.push((format!("file {}", page.id()), Box::new(move |app: &mut SheetApp| app.open_backstage(page))));
    }
    for name in [
        "about",
        "accessibility",
        "advancedFilter",
        "agents",
        "allowEditRanges",
        "analysisToolPak",
        "arrangeAll",
        "assignAction",
        "cfQuick",
        "chartTitle",
        "columnWidth",
        "commandSearch",
        "comment",
        "comments",
        "defineName",
        "deleteCells",
        "equation",
        "errorChecking",
        "evaluateFormula",
        "find",
        "forecastSheet",
        "formatCells",
        "goTo",
        "goToSpecial",
        "insertCells",
        "insertFunction",
        "insertLink",
        "insertPaste",
        "insertPicture",
        "journal",
        "macroEditor",
        "macros",
        "manageRules",
        "moveSheet",
        "nameManager",
        "open",
        "openLink",
        "openPassword",
        "outlineSettings",
        "pageSetup",
        "pasteName",
        "pasteSpecial",
        "protectSheet",
        "protectWorkbook",
        "recommendedCharts",
        "recordMacro",
        "removeDuplicates",
        "rowHeight",
        "saveAs",
        "saveChanges",
        "saveCopy",
        "selectData",
        "series",
        "solver",
        "sort",
        "sparkline",
        "spelling",
        "start",
        "statistics",
        "subtotal",
        "tableResize",
        "textToColumns",
        "unprotectWorkbook",
        "vbaEditor",
        "zoom",
    ] {
        v.push((format!("dialog {name}"), Box::new(move |app: &mut SheetApp| app.open_dialog(name, json!({})))));
    }
    v
}

/// Labels drawn in English in a Spanish interface although the Spanish catalog translates them.
fn untranslated(setup: &dyn Fn(&mut SheetApp)) -> Vec<String> {
    same_in_both(setup).into_iter().filter(|label| Language::Es.tr(label) != label.as_str()).collect()
}

/// A screen that draws a label the catalog translates, without going through the catalog, shows
/// English to everyone who runs Celdas in Spanish.
#[test]
fn spanish_interface_has_no_untranslated_label() {
    let mut failures = Vec::new();
    for (name, setup) in surfaces() {
        let labels = untranslated(setup.as_ref());
        if !labels.is_empty() {
            failures.push(format!("{name}: {labels:?}"));
        }
    }
    assert!(failures.is_empty(), "labels drawn in English although Spanish has them:\n{}", failures.join("\n"));
}

/// Prints every surface's untranslated labels: `cargo test -p celdas-ui-egui --test it -- --ignored --nocapture untranslated_report`.
#[test]
#[ignore = "a report, not a check"]
fn untranslated_report() {
    for (name, setup) in surfaces() {
        let same = same_in_both(setup.as_ref());
        println!("== {name}");
        for label in same {
            println!("   {label:?}");
        }
    }
}

/// Every command label and menu path part the engine registers that the Spanish catalog leaves as
/// it is (command search and the ribbon show them).
fn untranslated_commands() -> BTreeSet<String> {
    let session = crate::support::blank_session();
    let mut out = BTreeSet::new();
    for c in session.commands() {
        for text in std::iter::once(c.label).chain(c.menu.iter().copied()) {
            if Language::Es.tr(text) == text && !text.is_empty() {
                out.insert(text.to_owned());
            }
        }
    }
    out
}

/// Words that are the same in Spanish, and the product names the production sources must not spell out
/// (`display_names.rs`).
const SAME_IN_SPANISH: &[&str] = &["100%", "Format WordArt", "Macros", "Normal", "Series…", "Solver…", "Subtotal", "WordArt", "Zoom"];

/// A command shown in command search or the ribbon with a label the Spanish catalog lacks is shown
/// in English to everyone who runs Celdas in Spanish.
#[test]
fn every_command_label_and_menu_is_translated_to_spanish() {
    let missing: Vec<String> = untranslated_commands().into_iter().filter(|t| !SAME_IN_SPANISH.contains(&t.as_str())).collect();
    assert!(missing.is_empty(), "command labels and menu names without a Spanish row: {missing:?}");
}

#[test]
#[ignore = "a report, not a check"]
fn untranslated_command_report() {
    for text in untranslated_commands() {
        println!("   {text:?}");
    }
}
