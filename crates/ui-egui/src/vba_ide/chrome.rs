//! The editor's frame: the keys of the Visual Basic editor, the toolbar, the project explorer
//! with its outline, the module tabs, the panel and status line, the welcome state of a workbook
//! with no macros, and the questions (remove a module, save changes).

use egui::{Color32, Context, Frame, Id, Key, Margin, Modifiers, RichText, Sense, Ui, vec2};
use serde_json::json;

use super::{Adding, Ask, Bottom, Ide, Module, Severity, editor, model, panels};
use crate::SheetApp;
use crate::icons::{self, Icon};
use crate::theme::{self, Tokens};

/// What a new project starts with.
pub(super) const STARTER: &str =
    "Option Explicit\r\n\r\nSub MyMacro()\r\n    ' Your code goes here.\r\n    Range(\"A1\").Value = \"Hello from Celdas\"\r\nEnd Sub\r\n";

/// The keys of the Visual Basic editor.
pub fn keys(app: &mut SheetApp, ctx: &Context, ide: &mut Ide, leave: &mut bool) {
    // A dialog, a message or a question over the editor has the keyboard.
    if app.dialog.is_some() || app.message.is_some() || app.macros.question.is_some() || ide.ask.is_some() {
        return;
    }
    let hit = |mods: Modifiers, key: Key| ctx.input_mut(|i| i.consume_key(mods, key));
    let shift = Modifiers::SHIFT;
    let cmd = Modifiers::COMMAND;
    if hit(Modifiers::ALT, Key::F11) {
        close(ide, leave);
        return;
    }
    if hit(cmd, Key::S) {
        ide.save_all(app);
    }
    if hit(cmd, Key::F) {
        open_find(ide, false);
    }
    if hit(cmd, Key::H) {
        open_find(ide, true);
    }
    if hit(Modifiers::NONE, Key::F3) {
        editor::find_step(ide, true);
    }
    if hit(shift, Key::F3) {
        editor::find_step(ide, false);
    }
    if hit(cmd | shift, Key::F) {
        editor::format(app, ide);
    }
    if hit(cmd, Key::J) {
        ide.bottom_open = !ide.bottom_open;
    }
    if hit(Modifiers::NONE, Key::F12) {
        editor::go_to_definition(app, ide);
    }
    if hit(Modifiers::NONE, Key::F9) {
        let line = model::line_col(&ide.text_of(&ide.current.clone()), ide.selection.0).0;
        ide.toggle_breakpoint(app, line);
    }
    if hit(Modifiers::NONE, Key::F5) {
        run_or_continue(app, ide);
    }
    if hit(shift, Key::F5) {
        ide.stop(app, true);
    }
    if hit(Modifiers::NONE, Key::F8) {
        step_into(app, ide);
    }
    if hit(shift, Key::F8) {
        ide.step(app, "over");
    }
    if hit(cmd | shift, Key::F8) {
        ide.step(app, "out");
    }
    if ide.find.open && ide.popup.is_none() && hit(Modifiers::NONE, Key::Escape) {
        ide.find.open = false;
        ide.focus_editor = true;
    }
}

fn open_find(ide: &mut Ide, replace: bool) {
    ide.find.open = true;
    ide.find.replace |= replace;
    ide.find.focus = true;
    // The word under the caret (or the selection) is what a search most often starts from.
    let text = ide.text_of(&ide.current.clone());
    let (a, b) = (ide.selection.0.min(ide.selection.1), ide.selection.0.max(ide.selection.1));
    let picked: String = if a != b && !text.chars().skip(a).take(b - a).any(|c| c == '\n') {
        text.chars().skip(a).take(b - a).collect()
    } else {
        model::word_at(&text, ide.selection.0).map(|(_, w)| w).unwrap_or_default()
    };
    if !picked.is_empty() {
        ide.find.query = picked;
        ide.find.index = 0;
    }
}

fn close(ide: &mut Ide, leave: &mut bool) {
    if ide.has_unsaved_changes() {
        ide.ask = Some(Ask::Close);
    } else {
        *leave = true;
    }
}

fn run_or_continue(app: &mut SheetApp, ide: &mut Ide) {
    if ide.paused() {
        ide.step(app, "continue");
    } else if let Some(p) = ide.target() {
        ide.run(app, &p);
    } else {
        ide.status = Some(app.ui.language.tr("Put the caret inside a Sub that takes no arguments to run it.").to_string());
    }
}

fn step_into(app: &mut SheetApp, ide: &mut Ide) {
    if ide.paused() {
        ide.step(app, "into");
    } else if let Some(p) = ide.target() {
        ide.debug(app, &p);
    } else {
        ide.status = Some(app.ui.language.tr("Put the caret inside a Sub that takes no arguments to debug it.").to_string());
    }
}

fn keys_text(ui: &Ui, keys: &str) -> String {
    if ui.ctx().os() == egui::os::OperatingSystem::Mac { keys.replace("Ctrl+", "⌘").replace("Shift+", "⇧") } else { keys.to_string() }
}

/// A toolbar button. True when it was clicked while enabled.
fn tool(ui: &mut Ui, icon: Icon, enabled: bool, color: Color32, tip: &str, keys: &str) -> bool {
    let t = Tokens::get(ui.ctx());
    let (rect, resp) = ui.allocate_exact_size(vec2(32.0, 32.0), Sense::click());
    if enabled && resp.is_pointer_button_down_on() {
        ui.painter().rect_filled(rect, 6.0, t.pressed);
    } else if enabled && resp.hovered() {
        ui.painter().rect_filled(rect, 6.0, t.hover);
    }
    icons::paint(ui.painter(), egui::Rect::from_center_size(rect.center(), vec2(18.0, 18.0)), icon, if enabled { color } else { t.text_disabled });
    let name = crate::i18n::ui_tr(ui, tip).to_string();
    resp.widget_info(|| egui::WidgetInfo::labeled(egui::Role::Button, enabled, &name));
    let tip = if keys.is_empty() { name } else { format!("{name} ({})", keys_text(ui, keys)) };
    enabled && resp.on_hover_text(tip).clicked()
}

fn divider(ui: &mut Ui, t: &Tokens) {
    let (rect, _) = ui.allocate_exact_size(vec2(9.0, 32.0), Sense::hover());
    ui.painter().line_segment([rect.center_top() + vec2(0.0, 8.0), rect.center_bottom() - vec2(0.0, 8.0)], egui::Stroke::new(1.0, t.separator));
}

/// The whole editor.
pub fn layout(app: &mut SheetApp, ui: &mut Ui, ide: &mut Ide, leave: &mut bool) {
    let t = Tokens::get(ui.ctx());
    let g = theme::ISLAND_GAP as i8;
    status_line(app, ui, ide, &t);
    egui::Panel::top("vba_ide_toolbar")
        .exact_size(46.0 + f32::from(g))
        .show_separator_line(false)
        .frame(theme::island(&t, Margin::symmetric(10, 6), Margin { left: g, right: g, top: 0, bottom: g }))
        .show(ui, |ui| toolbar(app, ui, ide, leave, &t));
    if !ide.project {
        egui::CentralPanel::default().frame(Frame::NONE.fill(t.input_bg).outer_margin(Margin { left: g, right: g, top: 0, bottom: g })).show(
            ui,
            |ui| {
                welcome(app, ui, ide, &t);
                theme::round_island_corners(ui.painter(), &t, ui.max_rect(), true);
            },
        );
        return;
    }
    egui::Panel::left("vba_ide_explorer")
        .resizable(true)
        .default_size(250.0)
        .size_range(190.0..=420.0)
        .show_separator_line(false)
        .frame(theme::island(&t, Margin::symmetric(8, 8), Margin { left: g, right: 0, top: 0, bottom: g }))
        .show(ui, |ui| explorer(app, ui, ide, &t));
    if ide.bottom_open {
        egui::Panel::bottom("vba_ide_panel")
            .resizable(true)
            .default_size(220.0)
            .size_range(110.0..=560.0)
            .show_separator_line(false)
            .frame(theme::island(&t, Margin { left: 10, right: 10, top: 2, bottom: 8 }, Margin { left: g, right: g, top: 0, bottom: g }))
            .show(ui, |ui| panels::show(app, ui, ide));
    }
    let below = if ide.bottom_open { 0 } else { g };
    egui::CentralPanel::default().frame(Frame::NONE.fill(t.input_bg).outer_margin(Margin { left: g, right: g, top: 0, bottom: below })).show(
        ui,
        |ui| {
            tabs(app, ui, ide, &t);
            if ide.current.is_empty() {
                ui.add_space(24.0);
                ui.vertical_centered(|ui| {
                    ui.label(RichText::new(crate::i18n::ui_tr(ui, "Pick a module on the left to edit its code.")).color(t.text_dim));
                });
            } else {
                editor::show(app, ui, ide);
            }
            theme::round_island_corners(ui.painter(), &t, ui.max_rect(), below > 0);
        },
    );
}

fn toolbar(app: &mut SheetApp, ui: &mut Ui, ide: &mut Ide, leave: &mut bool, t: &Tokens) {
    let lang = app.ui.language;
    let paused = ide.paused();
    let target = ide.target();
    ui.horizontal_centered(|ui| {
        ui.spacing_mut().item_spacing.x = 2.0;
        if tool(ui, Icon::Back, true, t.text, "Back to the workbook", "Alt+F11") {
            close(ide, leave);
        }
        ui.add_space(6.0);
        let book = app.session.active().map(|d| d.display_title()).unwrap_or_default();
        ui.label(RichText::new(lang.tr("Visual Basic")).font(theme::ui_font(15.0)).strong().color(t.text));
        ui.add_space(8.0);
        ui.label(RichText::new(book).font(theme::ui_font(12.5)).color(t.text_dim));
        ui.add_space(6.0);
        divider(ui, t);
        if tool(ui, Icon::Save, ide.has_unsaved_changes(), t.text, "Save", "Ctrl+S") {
            ide.save_all(app);
        }
        if tool(ui, Icon::Code, ide.project && !paused, t.text, "Format Module", "Ctrl+Shift+F") {
            editor::format(app, ide);
        }
        if tool(ui, Icon::Search, ide.project, t.text, "Find and Replace", "Ctrl+F") {
            open_find(ide, false);
        }
        divider(ui, t);
        let can_run = paused || target.is_some();
        if paused {
            if tool(ui, Icon::Continue, true, t.accent, "Continue", "F5") {
                ide.step(app, "continue");
            }
        } else if tool(ui, Icon::Play, ide.project, if can_run { t.accent } else { t.text_dim }, "Run", "F5") {
            run_or_continue(app, ide);
        }
        if tool(ui, Icon::Step, paused || can_run, t.text, if paused { "Step Into" } else { "Start Debugging" }, "F8") {
            step_into(app, ide);
        }
        if tool(ui, Icon::StepOver, paused, t.text, "Step Over", "Shift+F8") {
            ide.step(app, "over");
        }
        if tool(ui, Icon::StepOut, paused, t.text, "Step Out", "Ctrl+Shift+F8") {
            ide.step(app, "out");
        }
        if tool(ui, Icon::StopSquare, paused, t.danger, "Stop", "Shift+F5") {
            ide.stop(app, true);
        }
        ui.with_layout(egui::Layout::right_to_left(egui::Align::Center), |ui| {
            if tool(ui, Icon::Terminal, ide.project, if ide.bottom_open { t.accent } else { t.text }, "Show or Hide the Panel", "Ctrl+J") {
                ide.bottom_open = !ide.bottom_open;
            }
            if paused {
                let procedure = ide.debug.as_ref().and_then(|d| d["procedure"].as_str()).unwrap_or("").to_string();
                ui.label(RichText::new(lang.tr_fill("Debugging “{name}”", &[("name", &procedure)])).font(theme::ui_font(12.5)).color(t.warning));
            }
        });
    });
}

/// The status line: what is going on, the problems, where the caret is.
fn status_line(app: &mut SheetApp, ui: &mut Ui, ide: &mut Ide, t: &Tokens) {
    let lang = app.ui.language;
    egui::Panel::bottom("vba_ide_status")
        .exact_size(28.0)
        .show_separator_line(false)
        .frame(Frame::NONE.fill(t.status_bar).inner_margin(Margin::symmetric(14, 0)))
        .show(ui, |ui| {
            ui.horizontal_centered(|ui| {
                let small = |text: String, color: Color32| RichText::new(text).font(theme::ui_font(12.5)).color(color);
                if let Some(f) = &ide.failure {
                    ui.label(small(f.message.lines().next().unwrap_or("").to_string(), t.danger));
                } else if let Some((m, l)) = ide.paused_at() {
                    let (rect, _) = ui.allocate_exact_size(vec2(10.0, 10.0), Sense::hover());
                    ui.painter().circle_filled(rect.center(), 4.5, t.breakpoint);
                    ui.label(small(lang.tr_fill("Paused at {module}, line {line}", &[("module", &m), ("line", &l.to_string())]), t.text));
                } else if let Some(s) = &ide.status {
                    ui.label(small(s.lines().next().unwrap_or("").to_string(), t.text_dim));
                } else {
                    ui.label(small(lang.tr("Ready").to_string(), t.text_dim));
                }
                let (errors, warnings) = ide.problems().iter().fold((0, 0), |(e, w), p| match p.2 {
                    Severity::Error => (e + 1, w),
                    Severity::Warning => (e, w + 1),
                    Severity::Hint => (e, w),
                });
                if errors + warnings > 0 {
                    ui.add_space(12.0);
                    let text =
                        lang.tr_fill("{errors} errors, {warnings} warnings", &[("errors", &errors.to_string()), ("warnings", &warnings.to_string())]);
                    if ui.add(egui::Button::new(small(text, if errors > 0 { t.danger } else { t.warning })).frame(false)).clicked() {
                        ide.bottom = Bottom::Problems;
                        ide.bottom_open = true;
                    }
                }
                ui.with_layout(egui::Layout::right_to_left(egui::Align::Center), |ui| {
                    if !ide.current.is_empty() {
                        let text = ide.text_of(&ide.current.clone());
                        let (line, col) = model::line_col(&text, ide.selection.0);
                        ui.label(small(
                            lang.tr_fill("Ln {line}, Col {col}", &[("line", &line.to_string()), ("col", &(col + 1).to_string())]),
                            t.text_dim,
                        ));
                        ui.add_space(12.0);
                        if ide.analysis.option_explicit {
                            ui.label(small("Option Explicit".to_string(), t.text_dim));
                        }
                    }
                });
            });
        });
}

/// A workbook with no VBA project.
fn welcome(app: &mut SheetApp, ui: &mut Ui, ide: &mut Ide, t: &Tokens) {
    let lang = app.ui.language;
    ui.add_space((ui.available_height() * 0.22).max(24.0));
    ui.vertical_centered(|ui| {
        let (rect, _) = ui.allocate_exact_size(vec2(56.0, 56.0), Sense::hover());
        ui.painter().rect_filled(rect, 14.0, t.accent_soft);
        icons::paint(ui.painter(), egui::Rect::from_center_size(rect.center(), vec2(30.0, 30.0)), Icon::Code, t.accent_text);
        ui.add_space(12.0);
        ui.label(RichText::new(lang.tr("Write macros in Visual Basic")).font(theme::ui_font(20.0)).strong().color(t.text));
        ui.add_space(4.0);
        ui.add(egui::Label::new(RichText::new(lang.tr("Create a macro project to write procedures that work on your sheets, with suggestions as you type, breakpoints and step-by-step debugging.")).color(t.text_dim)).wrap());
        ui.add_space(16.0);
        ui.horizontal(|ui| {
            ui.add_space((ui.available_width() - 330.0).max(0.0) / 2.0);
            if ui.add(egui::Button::new(RichText::new(lang.tr("Create Macro Project")).color(t.on_accent)).fill(t.accent).min_size(vec2(170.0, 34.0))).clicked() {
                match app.run("vba.create", json!({"code": STARTER})) {
                    Ok(_) => {
                        ide.reload(app);
                        ide.go_to("Module1", 4, 0);
                    }
                    Err(e) => app.message = Some(("Visual Basic".into(), crate::clean_error(&e))),
                }
            }
            if ui.add(egui::Button::new(lang.tr("Record a Macro")).min_size(vec2(140.0, 34.0))).clicked() {
                crate::macro_ui::toggle_recording(app);
            }
        });
    });
}

fn module_icon(m: &Module) -> Icon {
    if m.document {
        if m.name.eq_ignore_ascii_case("ThisWorkbook") { Icon::Book } else { Icon::Sheet }
    } else if m.class {
        Icon::Shapes
    } else {
        Icon::Code
    }
}

/// The project explorer: modules, the New Module form and the outline of the module open.
fn explorer(app: &mut SheetApp, ui: &mut Ui, ide: &mut Ide, t: &Tokens) {
    let lang = app.ui.language;
    ui.horizontal(|ui| {
        ui.label(RichText::new(lang.tr("PROJECT")).font(theme::ui_font(11.5)).strong().color(t.text_dim));
        ui.with_layout(egui::Layout::right_to_left(egui::Align::Center), |ui| {
            if crate::widgets::icon_button(ui, Icon::Plus, t.text, "New Module", vec2(26.0, 26.0)).clicked() {
                ide.adding = if ide.adding.is_some() { None } else { Some(Adding::default()) };
            }
        });
    });
    if ide.adding.is_some() {
        new_module(app, ui, ide, t);
    }
    egui::ScrollArea::vertical().auto_shrink([false, false]).show(ui, |ui| {
        ui.spacing_mut().item_spacing.y = 2.0;
        let mods = ide.mods.clone();
        let groups: [(&str, Vec<&Module>); 3] = [
            ("Standard Modules", mods.iter().filter(|m| !m.document && !m.class).collect()),
            ("Class Modules", mods.iter().filter(|m| !m.document && m.class).collect()),
            ("Workbook and Sheets", mods.iter().filter(|m| m.document).collect()),
        ];
        for (title, group) in groups {
            if group.is_empty() {
                continue;
            }
            ui.add_space(6.0);
            ui.label(RichText::new(lang.tr(title)).font(theme::ui_font(11.5)).color(t.text_dim));
            for m in group {
                module_row(ui, ide, m, t, lang);
            }
        }
        if mods.is_empty() {
            ui.add_space(8.0);
            ui.label(RichText::new(lang.tr("The project has no modules.")).color(t.text_dim));
        }
        outline(app, ui, ide, t);
    });
}

fn module_row(ui: &mut Ui, ide: &mut Ide, m: &Module, t: &Tokens, lang: crate::i18n::Language) {
    let selected = m.name.eq_ignore_ascii_case(&ide.current);
    let (rect, resp) = ui.allocate_exact_size(vec2(ui.available_width(), 30.0), Sense::click());
    if selected {
        ui.painter().rect_filled(rect, 6.0, t.accent_soft);
    } else if resp.hovered() {
        ui.painter().rect_filled(rect, 6.0, t.hover);
    }
    icons::paint(
        ui.painter(),
        egui::Rect::from_center_size(rect.left_center() + vec2(16.0, 0.0), vec2(16.0, 16.0)),
        module_icon(m),
        if selected { t.accent_text } else { t.text_dim },
    );
    let name_font = theme::ui_font(13.5);
    ui.painter().text(rect.left_center() + vec2(34.0, 0.0), egui::Align2::LEFT_CENTER, &m.name, name_font, t.text);
    let dirty = ide.drafts.get(&m.name).is_some_and(|d| *d != m.code);
    let broken = ide.errors.iter().any(|(n, e)| n.eq_ignore_ascii_case(&m.name) && !e.is_empty());
    let mut right = rect.right() - 12.0;
    if !m.document && resp.hovered() {
        let x = egui::Rect::from_center_size(egui::pos2(right - 4.0, rect.center().y), vec2(22.0, 22.0));
        let over = ui.input(|i| i.pointer.hover_pos()).is_some_and(|p| x.contains(p));
        if over {
            ui.painter().rect_filled(x, 4.0, t.pressed);
        }
        icons::paint(
            ui.painter(),
            egui::Rect::from_center_size(x.center(), vec2(14.0, 14.0)),
            Icon::Delete,
            if over { t.danger } else { t.text_dim },
        );
        if over && resp.clicked() {
            ide.ask = Some(Ask::Remove(m.name.clone()));
            return;
        }
        right -= 26.0;
    }
    if broken {
        icons::paint(ui.painter(), egui::Rect::from_center_size(egui::pos2(right - 6.0, rect.center().y), vec2(14.0, 14.0)), Icon::Warn, t.warning);
        right -= 20.0;
    }
    if dirty {
        ui.painter().circle_filled(egui::pos2(right - 6.0, rect.center().y), 3.5, t.accent);
    }
    let tip = if m.document {
        "The code behind a sheet or the workbook"
    } else if m.class {
        "Class module"
    } else {
        "Module"
    };
    if resp.on_hover_text(lang.tr(tip)).clicked() {
        ide.open_module(&m.name);
    }
}

fn new_module(app: &mut SheetApp, ui: &mut Ui, ide: &mut Ide, t: &Tokens) {
    let lang = app.ui.language;
    let Some(mut form) = ide.adding.take() else { return };
    let mut done = false;
    let mut cancel = false;
    Frame::new().fill(t.window).corner_radius(8.0).inner_margin(8).show(ui, |ui| {
        let field = ui.add(
            egui::TextEdit::singleline(&mut form.name)
                .hint_text(lang.tr("Module name"))
                .desired_width(f32::INFINITY)
                .id(Id::unique("celdas.vba_ide.new_module")),
        );
        if !field.has_focus() && form.name.is_empty() {
            field.request_focus();
        }
        ui.horizontal(|ui| {
            ui.selectable_value(&mut form.class, false, lang.tr("Module"));
            ui.selectable_value(&mut form.class, true, lang.tr("Class"));
        });
        let name = form.name.trim();
        let clash = ide.mods.iter().any(|m| m.name.eq_ignore_ascii_case(name));
        let valid = model::valid_module_name(name) && !clash;
        if !name.is_empty() && !valid {
            let why = if clash { "A module with that name exists." } else { "Use letters, digits and _, starting with a letter." };
            ui.label(RichText::new(lang.tr(why)).font(theme::ui_font(12.0)).color(t.danger));
        }
        ui.horizontal(|ui| {
            let enter = field.lost_focus() && ui.input(|i| i.key_pressed(Key::Enter));
            if ui.add_enabled(valid, egui::Button::new(lang.tr("Add"))).clicked() || (enter && valid) {
                done = true;
            }
            if ui.button(lang.tr("Cancel")).clicked() {
                cancel = true;
            }
        });
    });
    if done {
        let name = form.name.trim().to_string();
        let kind = if form.class { "class" } else { "standard" };
        match app.run("vba.save", json!({"module": name, "code": "Option Explicit\r\n", "kind": kind})) {
            Ok(_) => {
                ide.reload(app);
                ide.open_module(&name);
                ide.status = Some(lang.tr("Module added.").to_string());
            }
            Err(e) => app.message = Some(("Visual Basic".into(), crate::clean_error(&e))),
        }
    } else if !cancel {
        ide.adding = Some(form);
    }
}

/// The procedures of the module open; a Sub that can be started by hand has a Run button.
fn outline(app: &mut SheetApp, ui: &mut Ui, ide: &mut Ide, t: &Tokens) {
    let lang = app.ui.language;
    if ide.analysis.procs.is_empty() || ide.current.is_empty() {
        return;
    }
    ui.add_space(10.0);
    let arrow = if ide.outline_folded { Icon::Chevron } else { Icon::ChevronUp };
    ui.horizontal(|ui| {
        if crate::widgets::icon_button(ui, arrow, t.text_dim, "Outline", vec2(22.0, 22.0)).clicked() {
            ide.outline_folded = !ide.outline_folded;
        }
        ui.label(RichText::new(lang.tr("OUTLINE")).font(theme::ui_font(11.5)).strong().color(t.text_dim));
    });
    if ide.outline_folded {
        return;
    }
    let procs = ide.analysis.procs.clone();
    let here = ide.procedure_at_caret().map(|p| p.name.clone());
    let module = ide.current.clone();
    let mut go = None;
    let mut run = None;
    for p in &procs {
        let (rect, resp) = ui.allocate_exact_size(vec2(ui.available_width(), 28.0), Sense::click());
        let current = here.as_deref() == Some(p.name.as_str());
        if current || resp.hovered() {
            ui.painter().rect_filled(rect, 6.0, t.hover);
        }
        let color = if p.kind == "sub" {
            t.cat(theme::Cat::Blue)
        } else if p.kind == "function" {
            t.cat(theme::Cat::Green)
        } else {
            t.cat(theme::Cat::Orange)
        };
        let letter = match p.kind.as_str() {
            "sub" => "S",
            "function" => "F",
            _ => "P",
        };
        let badge = egui::Rect::from_center_size(rect.left_center() + vec2(14.0, 0.0), vec2(16.0, 16.0));
        ui.painter().rect_filled(badge, 4.0, color.gamma_multiply(0.22));
        ui.painter().text(badge.center(), egui::Align2::CENTER_CENTER, letter, theme::ui_font(10.5), color);
        let faded = if p.public { t.text } else { t.text_dim };
        ui.painter().text(rect.left_center() + vec2(30.0, 0.0), egui::Align2::LEFT_CENTER, &p.name, theme::ui_font(13.0), faded);
        if p.runnable() && resp.hovered() {
            let r = egui::Rect::from_center_size(egui::pos2(rect.right() - 16.0, rect.center().y), vec2(22.0, 22.0));
            let over = ui.input(|i| i.pointer.hover_pos()).is_some_and(|q| r.contains(q));
            if over {
                ui.painter().rect_filled(r, 4.0, t.pressed);
            }
            icons::paint(ui.painter(), egui::Rect::from_center_size(r.center(), vec2(14.0, 14.0)), Icon::Play, t.accent);
            if over && resp.clicked() {
                run = Some(p.name.clone());
                continue;
            }
        }
        if resp.on_hover_text(&p.signature).clicked() {
            go = Some((p.line, p.col));
        }
    }
    if let Some((line, col)) = go {
        ide.go_to(&module, line, col);
    }
    if let Some(name) = run {
        ide.run(app, &format!("{module}.{name}"));
    }
}

/// The module tabs above the code.
fn tabs(app: &mut SheetApp, ui: &mut Ui, ide: &mut Ide, t: &Tokens) {
    let lang = app.ui.language;
    let mut open = None;
    let mut shut = None;
    let strip = ui.available_width();
    let (bar, _) = ui.allocate_exact_size(vec2(strip, 36.0), Sense::hover());
    ui.painter().rect_filled(bar, 0.0, t.ribbon);
    let mut x = bar.left();
    let names = ide.tabs.clone();
    for name in &names {
        let active = name.eq_ignore_ascii_case(&ide.current);
        let dirty = ide.drafts.contains_key(name);
        let galley = ui.painter().layout_no_wrap(name.clone(), theme::ui_font(13.0), if active { t.text } else { t.text_dim });
        let w = galley.size().x + 52.0;
        let rect = egui::Rect::from_min_size(egui::pos2(x, bar.top() + 4.0), vec2(w, 32.0));
        let resp = ui.interact(rect, Id::unique("celdas.vba_ide.tab").with(name), Sense::click());
        if active {
            ui.painter().rect_filled(rect, egui::CornerRadius { nw: 8, ne: 8, sw: 0, se: 0 }, t.input_bg);
            ui.painter().line_segment([rect.left_top() + vec2(8.0, 0.5), rect.right_top() - vec2(8.0, -0.5)], egui::Stroke::new(2.0, t.accent));
        } else if resp.hovered() {
            ui.painter().rect_filled(rect.shrink2(vec2(2.0, 2.0)), 6.0, t.hover);
        }
        let pos = egui::pos2(rect.left() + 14.0, rect.center().y - galley.size().y / 2.0);
        ui.painter().galley(pos, galley, t.text);
        let x_rect = egui::Rect::from_center_size(egui::pos2(rect.right() - 15.0, rect.center().y), vec2(20.0, 20.0));
        let over_x = ui.input(|i| i.pointer.hover_pos()).is_some_and(|p| x_rect.contains(p));
        if dirty && !(resp.hovered() || active) {
            ui.painter().circle_filled(x_rect.center(), 3.5, t.accent);
        } else if resp.hovered() || active {
            if over_x {
                ui.painter().rect_filled(x_rect, 4.0, t.pressed);
            }
            icons::paint(ui.painter(), egui::Rect::from_center_size(x_rect.center(), vec2(12.0, 12.0)), Icon::Close, t.text_dim);
        }
        if resp.clicked() || resp.middle_clicked() {
            if over_x || resp.middle_clicked() {
                shut = Some(name.clone());
            } else {
                open = Some(name.clone());
            }
        }
        x += w + 2.0;
    }
    if ide.paused() {
        let text = lang.tr("Read-only while debugging");
        let galley = ui.painter().layout_no_wrap(text.to_string(), theme::ui_font(12.0), t.warning);
        let pos = egui::pos2(bar.right() - galley.size().x - 14.0, bar.center().y - galley.size().y / 2.0);
        ui.painter().galley(pos, galley, t.warning);
    }
    if let Some(n) = shut {
        ide.close_tab(&n);
    } else if let Some(n) = open {
        ide.open_module(&n);
    }
}

/// Remove Module? and Save changes?
pub fn questions(app: &mut SheetApp, ctx: &Context, ide: &mut Ide, leave: &mut bool) {
    if app.message.is_some() {
        return;
    }
    let lang = app.ui.language;
    match ide.ask.clone() {
        Some(Ask::Remove(name)) => {
            let mut answer = None;
            egui::Modal::new(Id::unique("celdas.vba_ide.remove")).show(ctx, |ui| {
                ui.set_width(380.0);
                ui.label(RichText::new(lang.tr_fill("Remove module “{name}”?", &[("name", &name)])).strong());
                ui.add_space(4.0);
                ui.label(lang.tr("This can be undone with Undo until the workbook is closed."));
                ui.add_space(10.0);
                ui.with_layout(egui::Layout::right_to_left(egui::Align::Center), |ui| {
                    if ui.button(lang.tr("Cancel")).clicked() || ui.input(|i| i.key_pressed(Key::Escape)) {
                        answer = Some(false);
                    }
                    if ui.add(egui::Button::new(lang.tr("Remove"))).clicked() {
                        answer = Some(true);
                    }
                });
            });
            match answer {
                Some(true) => {
                    ide.ask = None;
                    match app.run("vba.remove", json!({"module": name})) {
                        Ok(_) => {
                            ide.drafts.remove(&name);
                            ide.breakpoints.retain(|(m, _)| !m.eq_ignore_ascii_case(&name));
                            ide.close_tab(&name);
                            ide.reload(app);
                            ide.status = Some(lang.tr("Module removed.").to_string());
                        }
                        Err(e) => app.message = Some(("Visual Basic".into(), crate::clean_error(&e))),
                    }
                }
                Some(false) => ide.ask = None,
                None => {}
            }
        }
        Some(Ask::Close) => {
            let mut answer: Option<Option<bool>> = None;
            egui::Modal::new(Id::unique("celdas.vba_ide.close")).show(ctx, |ui| {
                ui.set_width(380.0);
                ui.label(RichText::new(lang.tr("Save changes to the VBA code?")).strong());
                ui.add_space(10.0);
                ui.with_layout(egui::Layout::right_to_left(egui::Align::Center), |ui| {
                    if ui.button(lang.tr("Cancel")).clicked() || ui.input(|i| i.key_pressed(Key::Escape)) {
                        answer = Some(None);
                    }
                    if ui.button(lang.tr("Don't Save")).clicked() {
                        answer = Some(Some(false));
                    }
                    if ui.add(egui::Button::new(lang.tr("Save"))).clicked() || ui.input(|i| i.key_pressed(Key::Enter)) {
                        answer = Some(Some(true));
                    }
                });
            });
            match answer {
                Some(Some(true)) => {
                    ide.ask = None;
                    if ide.save_all(app) {
                        *leave = true;
                    }
                }
                Some(Some(false)) => {
                    ide.ask = None;
                    *leave = true;
                }
                Some(None) => ide.ask = None,
                None => {}
            }
        }
        None => {}
    }
}
