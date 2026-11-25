//! The panel below the code: Problems, the Immediate window, Locals, Watch, the call stack and
//! the program's output.

use egui::{Id, Key, RichText, Sense, Ui, vec2};
use serde_json::Value as Json;

use super::{Bottom, Ide, Severity};
use crate::SheetApp;
use crate::code_editor;
use crate::icons::Icon;
use crate::theme::{self, Tokens};

const TABS: [(Bottom, &str); 6] = [
    (Bottom::Problems, "Problems"),
    (Bottom::Immediate, "Immediate"),
    (Bottom::Locals, "Locals"),
    (Bottom::Watch, "Watch"),
    (Bottom::Calls, "Call Stack"),
    (Bottom::Output, "Output"),
];

fn chip(ui: &mut Ui, label: &str, count: usize, active: bool, t: &Tokens) -> egui::Response {
    let font = theme::ui_font(12.5);
    let text = if count > 0 { format!("{label}  {count}") } else { label.to_string() };
    let galley = ui.painter().layout_no_wrap(text, font, if active { t.text } else { t.text_dim });
    let (rect, resp) = ui.allocate_exact_size(vec2(galley.size().x + 20.0, 28.0), Sense::click());
    if resp.hovered() && !active {
        ui.painter().rect_filled(rect.shrink2(vec2(2.0, 3.0)), 6.0, t.hover);
    }
    let pos = rect.center() - galley.size() / 2.0;
    ui.painter().galley(pos, galley, t.text);
    if active {
        let y = rect.bottom() - 1.0;
        ui.painter().line_segment([egui::pos2(rect.left() + 8.0, y), egui::pos2(rect.right() - 8.0, y)], egui::Stroke::new(2.0, t.accent));
    }
    resp.on_hover_cursor(egui::CursorIcon::PointingHand)
}

fn hint(ui: &mut Ui, text: &str, t: &Tokens) {
    ui.add_space(6.0);
    ui.label(RichText::new(text).font(theme::ui_font(12.5)).color(t.text_dim));
}

fn mono(text: impl Into<String>, color: egui::Color32) -> RichText {
    RichText::new(text).font(code_editor::font()).color(color)
}

/// The panel: the tab strip and what the selected tab holds.
pub fn show(app: &mut SheetApp, ui: &mut Ui, ide: &mut Ide) {
    let t = Tokens::get(ui.ctx());
    let lang = app.ui.language;
    let problems = ide.problems().len();
    ui.horizontal(|ui| {
        ui.spacing_mut().item_spacing.x = 0.0;
        for (tab, label) in TABS {
            let count = match tab {
                Bottom::Problems => problems,
                Bottom::Watch => ide.watches.len(),
                _ => 0,
            };
            if chip(ui, lang.tr(label), count, ide.bottom == tab, &t).clicked() {
                ide.bottom = tab;
                if tab == Bottom::Immediate {
                    ide.imm_focus = true;
                }
            }
        }
        ui.with_layout(egui::Layout::right_to_left(egui::Align::Center), |ui| {
            ui.add_space(6.0);
            if crate::widgets::icon_button(ui, Icon::Close, t.text_dim, lang.tr("Close panel"), vec2(24.0, 24.0)).clicked() {
                ide.bottom_open = false;
            }
            let clear = match ide.bottom {
                Bottom::Immediate => !ide.imm_log.is_empty(),
                Bottom::Output => !ide.output.is_empty(),
                _ => false,
            };
            if clear && ui.add(egui::Button::new(RichText::new(lang.tr("Clear")).font(theme::ui_font(12.0))).frame(false)).clicked() {
                if ide.bottom == Bottom::Immediate {
                    ide.imm_log.clear();
                } else {
                    ide.output.clear();
                }
            }
        });
    });
    ui.separator();
    match ide.bottom {
        Bottom::Problems => problems_tab(app, ui, ide, &t),
        Bottom::Immediate => immediate_tab(app, ui, ide, &t),
        Bottom::Locals => locals_tab(app, ui, ide, &t),
        Bottom::Watch => watch_tab(app, ui, ide, &t),
        Bottom::Calls => calls_tab(app, ui, ide, &t),
        Bottom::Output => output_tab(app, ui, ide, &t),
    }
}

fn problems_tab(app: &mut SheetApp, ui: &mut Ui, ide: &mut Ide, t: &Tokens) {
    let list = ide.problems();
    if list.is_empty() {
        hint(ui, app.ui.language.tr("No problems found."), t);
        return;
    }
    let mut go = None;
    egui::ScrollArea::vertical().auto_shrink([false, false]).show(ui, |ui| {
        for (module, line, severity, message) in &list {
            let (rect, resp) = ui.allocate_exact_size(vec2(ui.available_width(), 24.0), Sense::click());
            if resp.hovered() {
                ui.painter().rect_filled(rect, 4.0, t.hover);
            }
            let color = match severity {
                Severity::Error => t.danger,
                Severity::Warning => t.warning,
                Severity::Hint => t.text_dim,
            };
            ui.painter().circle_filled(rect.left_center() + vec2(12.0, 0.0), 4.0, color);
            let place = if *line > 0 { format!("{module}:{line}") } else { module.clone() };
            let width = ui.painter().layout_no_wrap(place.clone(), theme::ui_font(12.0), t.text_dim).size().x;
            ui.painter().text(rect.right_center() - vec2(10.0, 0.0), egui::Align2::RIGHT_CENTER, place, theme::ui_font(12.0), t.text_dim);
            let clip = ui.painter().with_clip_rect(egui::Rect::from_min_max(rect.min, egui::pos2(rect.right() - width - 20.0, rect.bottom())));
            clip.text(
                rect.left_center() + vec2(26.0, 0.0),
                egui::Align2::LEFT_CENTER,
                message.lines().next().unwrap_or(""),
                theme::ui_font(12.5),
                t.text,
            );
            if resp.clicked() {
                go = Some((module.clone(), *line));
            }
        }
    });
    if let Some((module, line)) = go {
        ide.go_to(&module, line.max(1), 0);
    }
}

fn immediate_tab(app: &mut SheetApp, ui: &mut Ui, ide: &mut Ide, t: &Tokens) {
    let lang = app.ui.language;
    let input_h = 30.0;
    let log_h = (ui.available_height() - input_h - 6.0).max(20.0);
    egui::ScrollArea::vertical().id_salt("vba_immediate_log").max_height(log_h).auto_shrink([false, false]).stick_to_bottom(true).show(ui, |ui| {
        if ide.imm_log.is_empty() {
            hint(ui, lang.tr("Type a statement, or ? followed by an expression, and press Enter."), t);
        }
        for line in &ide.imm_log {
            ui.label(mono(format!("› {}", line.input), t.text_dim));
            for out in &line.output {
                ui.label(mono(out.clone(), if line.error { t.danger } else { t.text }));
            }
        }
    });
    ui.horizontal(|ui| {
        let hint_text =
            if ide.paused() { lang.tr("Runs in the paused procedure") } else { lang.tr("Runs in this workbook, with the project's names") };
        let field = ui.add(
            egui::TextEdit::singleline(&mut ide.imm_input)
                .id(Id::unique("celdas.vba_ide.immediate"))
                .font(code_editor::font())
                .hint_text(hint_text)
                .desired_width(ui.available_width() - 70.0),
        );
        if ide.imm_focus {
            field.request_focus();
            ide.imm_focus = false;
        }
        if field.has_focus() {
            let recall = |ide: &mut Ide, up: bool| {
                let inputs: Vec<&str> = ide.imm_log.iter().map(|l| l.input.as_str()).collect();
                if inputs.is_empty() {
                    return;
                }
                let next = match (ide.imm_recall, up) {
                    (None, true) => Some(inputs.len() - 1),
                    (Some(i), true) => Some(i.saturating_sub(1)),
                    (Some(i), false) if i + 1 < inputs.len() => Some(i + 1),
                    _ => None,
                };
                ide.imm_input = next.and_then(|i| inputs.get(i)).map(|s| (*s).to_string()).unwrap_or_default();
                ide.imm_recall = next;
            };
            if ui.input_mut(|i| i.consume_key(egui::Modifiers::NONE, Key::ArrowUp)) {
                recall(ide, true);
            }
            if ui.input_mut(|i| i.consume_key(egui::Modifiers::NONE, Key::ArrowDown)) {
                recall(ide, false);
            }
        }
        let enter = field.lost_focus() && ui.input(|i| i.key_pressed(Key::Enter));
        let run = ui.add_enabled(!ide.imm_input.trim().is_empty(), egui::Button::new(lang.tr("Run")).min_size(vec2(56.0, 0.0))).clicked();
        if enter || run {
            ide.submit_immediate(app);
        }
    });
}

/// The frame whose variables are shown.
fn frame(ide: &Ide) -> Option<&Json> {
    ide.debug.as_ref()?["frames"].get(ide.frame)
}

fn variables_table(ui: &mut Ui, id: &str, list: &Json, t: &Tokens, lang: crate::i18n::Language) {
    let rows = list.as_array().map(Vec::as_slice).unwrap_or(&[]);
    egui::Grid::new(id).striped(true).num_columns(3).spacing(vec2(24.0, 4.0)).show(ui, |ui| {
        for h in ["Name", "Value", "Type"] {
            ui.label(RichText::new(lang.tr(h)).font(theme::ui_font(11.5)).color(t.text_dim));
        }
        ui.end_row();
        for v in rows {
            ui.label(mono(v["name"].as_str().unwrap_or(""), t.text));
            ui.label(mono(v["value"].as_str().unwrap_or(""), t.text));
            ui.label(RichText::new(v["type"].as_str().unwrap_or("")).font(theme::ui_font(12.0)).color(t.text_dim));
            ui.end_row();
        }
    });
}

fn not_paused(app: &SheetApp, ui: &mut Ui, t: &Tokens) {
    hint(ui, app.ui.language.tr("Start the macro with F8 to see its variables here."), t);
}

fn locals_tab(app: &mut SheetApp, ui: &mut Ui, ide: &mut Ide, t: &Tokens) {
    let lang = app.ui.language;
    let Some(f) = frame(ide).cloned() else {
        not_paused(app, ui, t);
        return;
    };
    let globals = ide.debug.as_ref().map(|d| d["globals"].clone()).unwrap_or(Json::Null);
    egui::ScrollArea::vertical().auto_shrink([false, false]).show(ui, |ui| {
        ui.label(
            RichText::new(format!("{} · {}", f["procedure"].as_str().unwrap_or(""), f["module"].as_str().unwrap_or("")))
                .font(theme::ui_font(12.5))
                .strong(),
        );
        if f["locals"].as_array().is_none_or(Vec::is_empty) {
            hint(ui, lang.tr("This procedure has no variables yet."), t);
        } else {
            variables_table(ui, "vba_locals", &f["locals"], t, lang);
        }
        if globals.as_array().is_some_and(|g| !g.is_empty()) {
            ui.add_space(8.0);
            ui.label(RichText::new(lang.tr("Module-level variables")).font(theme::ui_font(12.5)).strong());
            variables_table(ui, "vba_globals", &globals, t, lang);
        }
    });
}

fn watch_tab(app: &mut SheetApp, ui: &mut Ui, ide: &mut Ide, t: &Tokens) {
    let lang = app.ui.language;
    let mut remove = None;
    let mut add = false;
    ui.horizontal(|ui| {
        let field = ui.add(
            egui::TextEdit::singleline(&mut ide.watch_input)
                .id(Id::unique("celdas.vba_ide.watch"))
                .font(code_editor::font())
                .hint_text(lang.tr("Add a watch: a variable or an expression"))
                .desired_width(ui.available_width() - 70.0),
        );
        if field.lost_focus() && ui.input(|i| i.key_pressed(Key::Enter)) {
            add = true;
            field.request_focus();
        }
        if ui.add_enabled(!ide.watch_input.trim().is_empty(), egui::Button::new(lang.tr("Add")).min_size(vec2(56.0, 0.0))).clicked() {
            add = true;
        }
    });
    let values = ide.debug.as_ref().map(|d| d["watches"].clone()).unwrap_or(Json::Null);
    egui::ScrollArea::vertical().auto_shrink([false, false]).show(ui, |ui| {
        if ide.watches.is_empty() {
            hint(ui, lang.tr("Watches show the value of an expression at every step."), t);
        }
        egui::Grid::new("vba_watches").striped(true).num_columns(3).spacing(vec2(24.0, 4.0)).show(ui, |ui| {
            for (i, w) in ide.watches.iter().enumerate() {
                let found = values.as_array().and_then(|l| l.iter().find(|v| v["expression"].as_str() == Some(w.as_str())));
                ui.label(mono(w.clone(), t.text));
                match found {
                    Some(v) if v["error"].is_string() => ui.label(mono(v["error"].as_str().unwrap_or(""), t.danger)),
                    Some(v) => ui.label(mono(v["value"].as_str().unwrap_or(""), t.text)),
                    None => ui.label(mono("—", t.text_dim)),
                };
                if crate::widgets::icon_button(ui, Icon::Close, t.text_dim, lang.tr("Remove"), vec2(22.0, 22.0)).clicked() {
                    remove = Some(i);
                }
                ui.end_row();
            }
        });
    });
    if add {
        let w = ide.watch_input.trim().to_string();
        if !w.is_empty() && !ide.watches.contains(&w) && ide.watches.len() < 40 {
            ide.watches.push(w);
            ide.push_debug_inputs(app);
        }
        ide.watch_input.clear();
    }
    if let Some(i) = remove {
        ide.watches.remove(i);
        ide.push_debug_inputs(app);
    }
}

fn calls_tab(app: &mut SheetApp, ui: &mut Ui, ide: &mut Ide, t: &Tokens) {
    let Some(frames) = ide.debug.as_ref().and_then(|d| d["frames"].as_array()).cloned() else {
        not_paused(app, ui, t);
        return;
    };
    let mut pick = None;
    egui::ScrollArea::vertical().auto_shrink([false, false]).show(ui, |ui| {
        for (i, f) in frames.iter().enumerate() {
            let (rect, resp) = ui.allocate_exact_size(vec2(ui.available_width(), 24.0), Sense::click());
            if i == ide.frame {
                ui.painter().rect_filled(rect, 4.0, t.accent_soft);
            } else if resp.hovered() {
                ui.painter().rect_filled(rect, 4.0, t.hover);
            }
            let name = f["procedure"].as_str().unwrap_or("");
            let place = format!("{}:{}", f["module"].as_str().unwrap_or(""), f["line"].as_u64().unwrap_or(0));
            ui.painter().text(rect.left_center() + vec2(10.0, 0.0), egui::Align2::LEFT_CENTER, name, code_editor::font(), t.text);
            ui.painter().text(rect.right_center() - vec2(10.0, 0.0), egui::Align2::RIGHT_CENTER, place, theme::ui_font(12.0), t.text_dim);
            if resp.clicked() {
                pick = Some(i);
            }
        }
    });
    if let Some(i) = pick {
        ide.frame = i;
        if let Some((m, l)) = frames.get(i).and_then(|f| Some((f["module"].as_str()?.to_string(), u32::try_from(f["line"].as_u64()?).ok()?))) {
            ide.go_to(&m, l, 0);
        }
    }
}

fn output_tab(app: &mut SheetApp, ui: &mut Ui, ide: &mut Ide, t: &Tokens) {
    if ide.output.is_empty() {
        hint(ui, app.ui.language.tr("Nothing printed yet. Debug.Print writes here."), t);
        return;
    }
    egui::ScrollArea::vertical().auto_shrink([false, false]).stick_to_bottom(true).show(ui, |ui| {
        for line in &ide.output {
            ui.label(mono(line.clone(), t.text));
        }
    });
}
