//! The VBA editor (Automate › Visual Basic, Alt+F11, Macros › Edit on a VBA macro): the modules of
//! the workbook's VBA project and the code of the selected one, which can be edited, saved back
//! into the project, run, added and removed.
//!
//! The dialog keeps its state in the dialog's `values` and `result`:
//! - `result` is the project as the engine last reported it (`{source: vba.source, list: vba.list}`);
//! - `values.module` is the selected module, `values.drafts` the edited code by module name (a draft
//!   equal to the saved code is dropped), `values.status` the last message.
//!
//! Everything goes through the engine's `vba.save`, `vba.remove` and `vba.run` commands, so the
//! editor does nothing an agent can't do; each save is one undo step, and the file's project (its
//! `Attribute` lines, forms and everything Celdas does not read) is kept as it was.

use egui::{Id, Ui, vec2};
use serde_json::{Map, Value as Json, json};

use crate::code_editor::{self, Lang};
use crate::dialogs::Dialog;
use crate::theme::Tokens;
use crate::{SheetApp, macro_ui};

/// One module of the project, as the editor lists it.
#[derive(Clone, Debug, PartialEq)]
struct Module {
    name: String,
    class: bool,
    /// The code behind ThisWorkbook or a sheet: it can be edited but not removed.
    document: bool,
    /// The code a person edits (the module without its `Attribute` lines).
    code: String,
}

fn modules(d: &Dialog) -> Vec<Module> {
    d.result
        .as_ref()
        .and_then(|r| r["source"]["modules"].as_array())
        .into_iter()
        .flatten()
        .map(|m| Module {
            name: m["name"].as_str().unwrap_or("").to_string(),
            class: m["kind"] == "class",
            document: m["document"] == true,
            code: m["code"].as_str().unwrap_or("").to_string(),
        })
        .collect()
}

fn drafts(d: &Dialog) -> Map<String, Json> {
    d.values.get("drafts").and_then(Json::as_object).cloned().unwrap_or_default()
}

/// The names of the modules whose draft differs from what the project holds.
fn dirty_modules(d: &Dialog) -> Vec<String> {
    let drafts = drafts(d);
    modules(d).into_iter().filter(|m| drafts.get(&m.name).and_then(Json::as_str).is_some_and(|t| t != m.code)).map(|m| m.name).collect()
}

/// The editor holds code that is not in the workbook yet.
pub fn has_unsaved_changes(d: &Dialog) -> bool {
    !dirty_modules(d).is_empty()
}

/// Opens the editor: `{module}` selects a module, `{procedure}` ("Module.Name") its module.
pub fn open(app: &mut SheetApp, params: &Json) -> Option<Dialog> {
    let doc = app.session.active()?;
    if doc.wb.vba_project.is_none() {
        app.message =
            Some(("Visual Basic".into(), "This workbook has no VBA project. Open an .xlsm (or .xltm) workbook to see and edit its macros.".into()));
        return None;
    }
    let module = params["module"]
        .as_str()
        .map(str::to_string)
        .or_else(|| params["procedure"].as_str().and_then(|p| p.split_once('.')).map(|(m, _)| m.to_string()))
        .unwrap_or_default();
    let mut d = Dialog::custom("vbaEditor", "Visual Basic", json!({"module": module, "drafts": {}}));
    reload(app, &mut d);
    Some(d)
}

/// Reads the project again (after a save, an add or a removal).
fn reload(app: &mut SheetApp, d: &mut Dialog) {
    let source = app.session.run("vba.source", json!({}));
    let list = app.session.run("vba.list", json!({}));
    match (source, list) {
        (Ok(source), Ok(list)) => d.result = Some(json!({"source": source, "list": list})),
        (Err(e), _) | (_, Err(e)) => {
            d.result = None;
            d.values.insert("status".into(), json!(crate::clean_error(&e)));
        }
    }
}

fn set_status(d: &mut Dialog, text: impl Into<String>) {
    d.values.insert("status".into(), json!(text.into()));
}

/// Saves every module with a draft. Returns whether all of them were saved.
fn save_all(app: &mut SheetApp, d: &mut Dialog) -> bool {
    let names = dirty_modules(d);
    if names.is_empty() {
        return true;
    }
    let mut drafts = drafts(d);
    let mut notes = Vec::new();
    let mut ok = true;
    for name in &names {
        let Some(code) = drafts.get(name).and_then(Json::as_str).map(str::to_string) else { continue };
        match app.run("vba.save", json!({"module": name, "code": code})) {
            Ok(r) => {
                drafts.remove(name);
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
    d.values.insert("drafts".into(), Json::Object(drafts));
    reload(app, d);
    if notes.is_empty() {
        let n = names.len().to_string();
        set_status(d, app.ui.language.tr_fill("Saved the code of {n} module(s). The rest of the project was kept as it was.", &[("n", &n)]));
    } else {
        // The code is saved even when Celdas can't read it; say what it could not.
        set_status(d, notes.join("\n"));
    }
    ok
}

/// What the dialog's body draws. `open` turns false when it closes.
pub fn body(app: &mut SheetApp, ui: &mut Ui, d: &mut Dialog, open: &mut bool) {
    let t = Tokens::get(ui.ctx());
    if d.result.is_none() {
        reload(app, d);
    }
    let mods = modules(d);
    // The project went away (the workbook changed under the editor).
    if app.session.active().is_none_or(|doc| doc.wb.vba_project.is_none()) {
        ui.label(egui::RichText::new(crate::i18n::ui_tr(ui, "The active workbook has no VBA project.")).color(t.danger));
        if ui.button(crate::i18n::ui_tr(ui, "Close")).clicked() {
            *open = false;
        }
        return;
    }
    let mut selected = d.values.get("module").and_then(Json::as_str).unwrap_or("").to_string();
    if !mods.iter().any(|m| m.name.eq_ignore_ascii_case(&selected)) {
        selected = mods.first().map(|m| m.name.clone()).unwrap_or_default();
    }
    let drafts = drafts(d);
    let current = mods.iter().find(|m| m.name.eq_ignore_ascii_case(&selected)).cloned();

    ui.horizontal_top(|ui| {
        // The modules.
        ui.vertical(|ui| {
            ui.set_width(180.0);
            ui.label(egui::RichText::new(crate::i18n::ui_tr(ui, "Modules")).strong());
            egui::ScrollArea::vertical().id_salt("vba_modules").max_height(430.0).show(ui, |ui| {
                for m in &mods {
                    let dirty = drafts.get(&m.name).and_then(Json::as_str).is_some_and(|t| t != m.code);
                    let label = format!("{}{}", m.name, if dirty { " •" } else { "" });
                    let tip = if m.document {
                        crate::i18n::ui_tr(ui, "The code behind a sheet or the workbook")
                    } else if m.class {
                        crate::i18n::ui_tr(ui, "Class module")
                    } else {
                        crate::i18n::ui_tr(ui, "Module")
                    };
                    if ui.selectable_label(m.name == selected, label).on_hover_text(tip).clicked() {
                        selected = m.name.clone();
                    }
                }
            });
            ui.add_space(6.0);
            new_module(app, ui, d, &mods);
            let removable = current.as_ref().is_some_and(|m| !m.document);
            if ui.add_enabled(removable, egui::Button::new(crate::i18n::ui_tr(ui, "Remove Module"))).clicked()
                && let Some(m) = &current
            {
                d.values.insert("askRemove".into(), json!(m.name));
            }
        });
        ui.add_space(8.0);
        // The code.
        ui.vertical(|ui| {
            let Some(m) = &current else {
                ui.label(egui::RichText::new(crate::i18n::ui_tr(ui, "The project has no modules.")).color(t.text_dim));
                return;
            };
            let mut text = drafts.get(&m.name).and_then(Json::as_str).unwrap_or(&m.code).to_string();
            let before = text.clone();
            ui.set_min_width(520.0);
            code_editor::code_editor(ui, Id::unique("celdas.vba_code").with(&m.name), &mut text, Lang::Vba, 26);
            if text != before {
                let mut next = drafts.clone();
                if text == m.code {
                    next.remove(&m.name);
                } else {
                    next.insert(m.name.clone(), json!(text));
                }
                d.values.insert("drafts".into(), Json::Object(next));
            }
        });
    });
    d.values.insert("module".into(), json!(selected));
    ui.add_space(6.0);
    if let Some(status) = d.values.get("status").and_then(Json::as_str).filter(|s| !s.is_empty()) {
        ui.label(egui::RichText::new(crate::i18n::ui_tr(ui, status)).color(t.text_dim));
    }
    let problems: Vec<String> = d
        .result
        .as_ref()
        .and_then(|r| r["list"]["errors"].as_array())
        .into_iter()
        .flatten()
        .filter(|e| e["module"].as_str().is_some_and(|m| m.eq_ignore_ascii_case(&selected)))
        .filter_map(|e| e["message"].as_str().map(str::to_string))
        .collect();
    for p in problems.iter().take(3) {
        ui.label(egui::RichText::new(format!("⚠ {p}")).color(t.danger));
    }
    ui.add_space(4.0);
    actions(app, ui, d, &selected, open);
    questions(app, ui.ctx(), d, open);
}

/// The Add Module row: a name and Module / Class.
fn new_module(app: &mut SheetApp, ui: &mut Ui, d: &mut Dialog, mods: &[Module]) {
    let mut name = d.values.get("newName").and_then(Json::as_str).unwrap_or("").to_string();
    ui.add(egui::TextEdit::singleline(&mut name).hint_text(crate::i18n::ui_tr(ui, "New module name")).desired_width(170.0));
    d.values.insert("newName".into(), json!(name));
    ui.horizontal(|ui| {
        for (label, kind) in [("Add Module", "standard"), ("Add Class", "class")] {
            let clash = mods.iter().any(|m| m.name.eq_ignore_ascii_case(name.trim()));
            let valid = !name.trim().is_empty() && !clash && name.trim().chars().all(|c| c.is_alphanumeric() || c == '_');
            if ui.add_enabled(valid, egui::Button::new(crate::i18n::ui_tr(ui, label))).clicked() {
                let module = name.trim().to_string();
                match app.run("vba.save", json!({"module": module, "code": "Option Explicit\r\n", "kind": kind})) {
                    Ok(_) => {
                        d.values.insert("module".into(), json!(module));
                        d.values.insert("newName".into(), json!(""));
                        reload(app, d);
                        set_status(d, "Module added.");
                    }
                    Err(e) => app.message = Some(("Visual Basic".into(), crate::clean_error(&e))),
                }
            }
        }
    });
}

/// Save, Revert, Run and Close.
fn actions(app: &mut SheetApp, ui: &mut Ui, d: &mut Dialog, selected: &str, open: &mut bool) {
    let list_rows: Vec<(String, String)> = d
        .result
        .as_ref()
        .and_then(|r| r["list"]["procedures"].as_array())
        .into_iter()
        .flatten()
        .filter(|p| p["kind"] == "sub" && p["public"] == true && p["takesArguments"] == false)
        .map(|p| (p["module"].as_str().unwrap_or("").to_string(), p["name"].as_str().unwrap_or("").to_string()))
        .collect();
    let dirty = has_unsaved_changes(d);
    ui.horizontal(|ui| {
        if ui.add_enabled(dirty, egui::Button::new(crate::i18n::ui_tr(ui, "Save")).min_size(vec2(72.0, 0.0))).clicked() {
            save_all(app, d);
        }
        if ui.add_enabled(dirty, egui::Button::new(crate::i18n::ui_tr(ui, "Revert"))).clicked() {
            let mut drafts = drafts(d);
            drafts.remove(selected);
            d.values.insert("drafts".into(), Json::Object(drafts));
            set_status(d, "");
        }
        ui.separator();
        let mut procedure = d.values.get("procedure").and_then(Json::as_str).unwrap_or("").to_string();
        // Prefer a procedure of the selected module.
        if !list_rows.iter().any(|(m, n)| format!("{m}.{n}") == procedure) {
            procedure = list_rows
                .iter()
                .find(|(m, _)| m.eq_ignore_ascii_case(selected))
                .or(list_rows.first())
                .map(|(m, n)| format!("{m}.{n}"))
                .unwrap_or_default();
        }
        let name = app.ui.language.tr("Procedure").to_string();
        let shown = if procedure.is_empty() { "—".to_string() } else { procedure.clone() };
        let combo = nodalia_ui::ComboBox::from_id_salt("vba_procedure").width(220.0).selected_text(shown).show_ui(ui, |ui| {
            for (m, n) in &list_rows {
                let key = format!("{m}.{n}");
                ui.selectable_value(&mut procedure, key.clone(), key);
            }
        });
        combo.response.accessible_name(name);
        if ui.add_enabled(!procedure.is_empty(), egui::Button::new(crate::i18n::ui_tr(ui, "Run"))).clicked() && save_all(app, d) {
            macro_ui::call(app, "Run Macro", "vba.run", json!({"procedure": procedure}));
        }
        d.values.insert("procedure".into(), json!(procedure));
        ui.with_layout(egui::Layout::right_to_left(egui::Align::Center), |ui| {
            if ui.button(crate::i18n::ui_tr(ui, "Close")).clicked() {
                request_close(d, open);
            }
        });
    });
}

/// Closing asks first when there is unsaved code.
pub fn request_close(d: &mut Dialog, open: &mut bool) {
    if has_unsaved_changes(d) {
        d.values.insert("askClose".into(), json!(true));
    } else {
        *open = false;
    }
}

/// Remove Module? and Save changes?
fn questions(app: &mut SheetApp, ctx: &egui::Context, d: &mut Dialog, open: &mut bool) {
    if let Some(name) = d.values.get("askRemove").and_then(Json::as_str).map(str::to_string) {
        let mut answer = None;
        egui::Modal::new(Id::unique("celdas.vba_remove")).show(ctx, |ui| {
            ui.set_width(360.0);
            ui.label(egui::RichText::new(crate::i18n::ui_fill(ui, "Remove module “{name}”?", &[("name", &name)])).strong());
            ui.label(crate::i18n::ui_tr(ui, "This can be undone with Undo until the workbook is closed."));
            ui.add_space(8.0);
            ui.with_layout(egui::Layout::right_to_left(egui::Align::Center), |ui| {
                if ui.button(crate::i18n::ui_tr(ui, "Cancel")).clicked() || ui.input(|i| i.key_pressed(egui::Key::Escape)) {
                    answer = Some(false);
                }
                if ui.button(crate::i18n::ui_tr(ui, "Remove")).clicked() {
                    answer = Some(true);
                }
            });
        });
        if let Some(yes) = answer {
            {
                d.values.remove("askRemove");
                if yes {
                    match app.run("vba.remove", json!({"module": name})) {
                        Ok(_) => {
                            let mut drafts = drafts(d);
                            drafts.remove(&name);
                            d.values.insert("drafts".into(), Json::Object(drafts));
                            d.values.insert("module".into(), json!(""));
                            reload(app, d);
                            set_status(d, "Module removed.");
                        }
                        Err(e) => app.message = Some(("Visual Basic".into(), crate::clean_error(&e))),
                    }
                }
            }
        }
        return;
    }
    if d.values.get("askClose").and_then(Json::as_bool) == Some(true) {
        let mut answer: Option<Option<bool>> = None;
        egui::Modal::new(Id::unique("celdas.vba_close")).show(ctx, |ui| {
            ui.set_width(360.0);
            ui.label(egui::RichText::new(crate::i18n::ui_tr(ui, "Save changes to the VBA code?")).strong());
            ui.add_space(8.0);
            ui.with_layout(egui::Layout::right_to_left(egui::Align::Center), |ui| {
                if ui.button(crate::i18n::ui_tr(ui, "Cancel")).clicked() || ui.input(|i| i.key_pressed(egui::Key::Escape)) {
                    answer = Some(None);
                }
                if ui.button(crate::i18n::ui_tr(ui, "Don't Save")).clicked() {
                    answer = Some(Some(false));
                }
                if ui.button(crate::i18n::ui_tr(ui, "Save")).clicked() || ui.input(|i| i.key_pressed(egui::Key::Enter)) {
                    answer = Some(Some(true));
                }
            });
        });
        match answer {
            Some(Some(true)) => {
                d.values.remove("askClose");
                if save_all(app, d) {
                    *open = false;
                }
            }
            Some(Some(false)) => {
                d.values.remove("askClose");
                *open = false;
            }
            Some(None) => {
                d.values.remove("askClose");
            }
            None => {}
        }
    }
}

/// Escape or the window's close button: unsaved code asks first, and while a question is up the
/// key answers that question instead.
pub fn escape(d: &mut Dialog, open: &mut bool) {
    if d.values.get("askClose").is_none() && d.values.get("askRemove").is_none() {
        request_close(d, open);
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn dialog(code: &str, draft: Option<&str>) -> Dialog {
        let mut d = Dialog::custom("vbaEditor", "Visual Basic", json!({"drafts": {}}));
        d.result = Some(json!({"source": {"modules": [{"name": "Module1", "kind": "standard", "document": false, "code": code}]}, "list": {}}));
        if let Some(t) = draft {
            d.values.insert("drafts".into(), json!({"Module1": t}));
        }
        d
    }

    #[test]
    fn a_draft_equal_to_the_saved_code_is_not_a_change() {
        assert!(!has_unsaved_changes(&dialog("Sub A()\nEnd Sub\n", None)));
        assert!(!has_unsaved_changes(&dialog("Sub A()\nEnd Sub\n", Some("Sub A()\nEnd Sub\n"))));
        assert!(has_unsaved_changes(&dialog("Sub A()\nEnd Sub\n", Some("Sub B()\nEnd Sub\n"))));
    }

    #[test]
    fn closing_with_unsaved_code_asks_first() {
        let mut d = dialog("Sub A()\nEnd Sub\n", Some("x"));
        let mut open = true;
        request_close(&mut d, &mut open);
        assert!(open && d.values.get("askClose") == Some(&json!(true)));
        let mut clean = dialog("a", None);
        request_close(&mut clean, &mut open);
        assert!(!open);
    }
}
