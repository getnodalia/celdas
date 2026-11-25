//! The code area: the text with its margin (line numbers, breakpoints, the line being debugged),
//! squiggles under mistakes, the highlight of the name under the caret and of search hits,
//! suggestions as you type, the call's signature, hover help (with the value of a variable while a
//! macro is paused) and the find and replace bar.

use std::ops::Range;

use egui::{Align2, Area, Color32, Event, Frame, Id, Key, Order, RichText, Sense, Ui, vec2};
use serde_json::{Value as Json, json};

use super::{Card, Ide, Item, Popup, Severity, Signature, model};
use crate::SheetApp;
use crate::code_editor::{self, FieldOptions, Gutter, Lang, Mark, MarkKind};
use crate::theme::{self, Cat, Tokens};

/// The most text the analysis keeps up with while typing.
const LIVE_LIMIT: usize = 300_000;
/// Rows of the suggestion list shown before it scrolls.
const POPUP_ROWS: f32 = 8.0;
const ROW: f32 = 22.0;

fn severity_color(t: &Tokens, s: Severity) -> Color32 {
    match s {
        Severity::Error => t.danger,
        Severity::Warning => t.warning,
        Severity::Hint => t.text_dim,
    }
}

/// A byte range for a mark that is never empty: a mistake at the end of a line still shows.
fn mark_range(text: &str, line: u32, start: usize, end: usize) -> Option<Range<usize>> {
    let end = if end <= start { start + 1 } else { end };
    let r = model::byte_range(text, line, start, end);
    (r.start < r.end).then_some(r)
}

/// The breakpoints of the current module.
fn lines_of(ide: &Ide) -> std::collections::BTreeSet<u32> {
    ide.breakpoints.iter().filter(|(m, _)| m.eq_ignore_ascii_case(&ide.current)).map(|(_, l)| *l).collect()
}

/// Records `after` as the module's text and moves the breakpoints with the lines they were on.
fn apply_edit(ide: &mut Ide, before: &str, after: &str) {
    let name = ide.current.clone();
    let moved = model::shift_lines(before, after, &lines_of(ide));
    ide.breakpoints.retain(|(m, _)| !m.eq_ignore_ascii_case(&name));
    ide.breakpoints.extend(moved.into_iter().map(|l| (name.clone(), l)));
    ide.set_text(&name, after.to_string());
    if ide.failure.as_ref().is_some_and(|f| f.module.eq_ignore_ascii_case(&name)) {
        ide.failure = None;
    }
}

/// The frame the stack panel selected: `(module, line)`.
fn frame_site(ide: &Ide) -> Option<(String, u32)> {
    let f = ide.debug.as_ref()?["frames"].get(ide.frame)?;
    Some((f["module"].as_str()?.to_string(), u32::try_from(f["line"].as_u64()?).ok()?))
}

fn marks(ide: &Ide, text: &str, t: &Tokens, hits: &[Range<usize>]) -> Vec<Mark> {
    let mut out = Vec::new();
    for r in hits {
        let current = ide.find.index < hits.len() && hits.get(ide.find.index) == Some(r);
        let color = if current { t.warning.gamma_multiply(0.5) } else { t.accent.gamma_multiply(0.28) };
        out.push(Mark { range: r.clone(), color, kind: MarkKind::Background });
    }
    for r in &ide.occurrences {
        out.push(Mark { range: r.clone(), color: t.accent_soft, kind: MarkKind::Background });
    }
    for d in &ide.analysis.diags {
        if let Some(range) = mark_range(text, d.line, d.start, d.end) {
            out.push(Mark { range, color: severity_color(t, d.severity), kind: MarkKind::Underline });
        }
    }
    out
}

fn gutter(ide: &Ide, t: &Tokens) -> Gutter {
    let mut flags: Vec<(u32, Color32)> = Vec::new();
    for d in &ide.analysis.diags {
        if d.severity == Severity::Hint {
            continue;
        }
        let color = severity_color(t, d.severity);
        match flags.iter_mut().find(|(l, _)| *l == d.line) {
            Some(f) if d.severity == Severity::Error => f.1 = color,
            Some(_) => {}
            None => flags.push((d.line, color)),
        }
    }
    if let Some(f) = ide.failure.as_ref().filter(|f| f.line > 0 && f.module.eq_ignore_ascii_case(&ide.current)) {
        flags.retain(|(l, _)| *l != f.line);
        flags.push((f.line, t.danger));
    }
    let arrow = frame_site(ide).filter(|(m, _)| m.eq_ignore_ascii_case(&ide.current)).map(|(_, l)| l);
    Gutter { width: 58.0, breakpoints: lines_of(ide).into_iter().collect(), arrow, flags, background: t.ribbon, separator: t.separator }
}

fn line_fills(ide: &Ide, text: &str, t: &Tokens) -> Vec<(u32, Color32)> {
    let mut fills = Vec::new();
    let (caret_line, _) = model::line_col(text, ide.selection.0);
    if !ide.paused() {
        fills.push((caret_line, t.hover));
    }
    if let Some(f) = ide.failure.as_ref().filter(|f| f.line > 0 && f.module.eq_ignore_ascii_case(&ide.current)) {
        fills.push((f.line, t.danger.gamma_multiply(0.16)));
    }
    if let Some((m, l)) = frame_site(ide)
        && m.eq_ignore_ascii_case(&ide.current)
    {
        fills.push((l, t.warning.gamma_multiply(0.22)));
    }
    fills
}

/// Draws the code of the current module (and the bar of find and replace over it).
pub fn show(app: &mut SheetApp, ui: &mut Ui, ide: &mut Ide) {
    let t = Tokens::get(ui.ctx());
    let name = ide.current.clone();
    let mut text = ide.text_of(&name);
    if ide.find.open {
        find_bar(app, ui, ide, &mut text);
    }
    let before_analysis = text.clone();
    ide.analyse(app, &before_analysis);
    let id = Id::unique("celdas.vba_ide.code").with(&name);
    if ide.focus_editor {
        ui.memory_mut(|m| m.request_focus(id));
        ide.focus_editor = false;
    }
    let hits = if ide.find.open && !ide.find.query.is_empty() { model::find_all(&text, &ide.find.query, ide.find.how) } else { Vec::new() };
    let popup_open = ide.popup.is_some();
    let stop = move |e: &Event| match e {
        Event::Key { key: Key::Space, pressed: true, modifiers, .. } => modifiers.command,
        Event::Key { key, pressed: true, .. } if popup_open => matches!(key, Key::ArrowUp | Key::ArrowDown | Key::Enter | Key::Tab | Key::Escape),
        _ => false,
    };
    let row_h = ui.ctx().fonts_mut(|f| f.row_height(&code_editor::font()));
    let height = (ui.available_height() - 2.0).max(row_h * 4.0);
    let options = FieldOptions {
        width: f32::INFINITY,
        max_height: height,
        interactive: !ide.paused() && app.dialog.is_none() && ide.ask.is_none(),
        stop: &stop,
        keep_escape: popup_open,
        caret: ide.reveal.take(),
        name: "Code Editor",
        marks: marks(ide, &text, &t, &hits),
        gutter: Some(gutter(ide, &t)),
        line_fills: line_fills(ide, &text, &t),
        min_rows: (height / row_h).ceil() as usize,
        ..FieldOptions::new(Lang::Vba)
    };
    let before = text.clone();
    let out = code_editor::code_field(ui, id, &mut text, &options);
    let changed = text != before;
    if changed {
        apply_edit(ide, &before, &text);
    }
    if let Some(line) = out.gutter_click {
        ide.toggle_breakpoint(app, line);
    }
    let selection = out.selection;
    let moved = selection != ide.last_selection;
    ide.last_selection = selection;
    ide.selection = selection;
    let focused = out.response.has_focus();
    if out.response.clicked() && ui.input(|i| i.modifiers.command) {
        go_to_definition(app, ide);
    }

    // Suggestions: after typing, or when asked for with Ctrl+Space.
    let explicit = std::mem::take(&mut ide.explicit);
    let mut asked = explicit;
    let mut accepted = None;
    if let Some(Event::Key { key, .. }) = out.stop.as_ref() {
        match key {
            Key::Space => asked = true,
            Key::ArrowDown | Key::ArrowUp => {
                if let Some(p) = ide.popup.as_mut() {
                    let n = p.items.len().max(1);
                    p.selected = if *key == Key::ArrowDown { (p.selected + 1) % n } else { (p.selected + n - 1) % n };
                    ui.data_mut(|d| d.insert_temp(id.with("scroll_to_selected"), true));
                }
            }
            Key::Enter | Key::Tab => accepted = ide.popup.as_ref().map(|p| p.selected),
            Key::Escape => ide.popup = None,
            _ => {}
        }
    }
    let at = out.caret.map(|r| r.left_bottom()).unwrap_or_else(|| ide.popup.as_ref().map_or(egui::Pos2::ZERO, |p| p.at));
    if (changed || asked) && focused && text.len() <= LIVE_LIMIT {
        suggest(app, ide, &text, selection.0, asked, at);
    } else if moved && !changed {
        ide.popup = None;
        let (line, _) = model::line_col(&text, selection.0);
        if model::line_text(&text, line).contains('(') && text.len() <= LIVE_LIMIT {
            suggest(app, ide, &text, selection.0, false, at);
            ide.popup = None;
        } else {
            ide.signature = None;
        }
    }
    if !focused && !changed && ide.popup.is_some() {
        // The list stays while the pointer is over it (a click on a row takes the focus first).
        let over = ui.input(|i| i.pointer.hover_pos()).is_some_and(|p| popup_rect(ide).is_some_and(|r| r.contains(p)));
        if !over && !ui.input(|i| i.pointer.any_down()) {
            ide.popup = None;
        }
    }
    if let Some(i) = accepted.or_else(|| popup_ui(ui, ide, &t, id)) {
        accept(ide, &mut text, i);
    }
    if ide.popup.is_none() {
        signature_ui(ui, ide, &t, out.caret.map(|r| r.left_top()));
    }

    occurrences(app, ide, &text, selection);
    hover(app, ui, ide, &text, out.hover, &t);
}

/// What the list's rectangle is, for keeping it while it is pointed at.
fn popup_rect(ide: &Ide) -> Option<egui::Rect> {
    let p = ide.popup.as_ref()?;
    Some(egui::Rect::from_min_size(p.at, vec2(360.0, ROW * POPUP_ROWS.min(p.items.len() as f32) + 40.0)))
}

fn suggest(app: &mut SheetApp, ide: &mut Ide, text: &str, caret: usize, explicit: bool, at: egui::Pos2) {
    let (line, col) = model::line_col(text, caret);
    let params = json!({"code": text, "line": line, "col": col, "explicit": explicit, "module": ide.current});
    let Ok(r) = app.session.run("vba.complete", params) else {
        ide.popup = None;
        return;
    };
    let from = usize::try_from(r["from"].as_u64().unwrap_or(0)).unwrap_or(0);
    let items: Vec<Item> = r["items"]
        .as_array()
        .into_iter()
        .flatten()
        .map(|i| Item {
            label: i["label"].as_str().unwrap_or("").to_string(),
            kind: i["kind"].as_str().unwrap_or("").to_string(),
            detail: i["detail"].as_str().unwrap_or("").to_string(),
            doc: i["doc"].as_str().unwrap_or("").to_string(),
            insert: i["insert"].as_str().unwrap_or("").to_string(),
        })
        .collect();
    // A list whose only entry is the word already typed would only be in the way.
    let typed: String = text.chars().skip(model::char_at(text, line, from)).take(col.saturating_sub(from)).collect();
    let useless = items.len() == 1 && items.first().is_some_and(|i| i.label.eq_ignore_ascii_case(&typed) && !explicit);
    ide.popup = (!items.is_empty() && !useless).then_some(Popup { from, items, selected: 0, at });
    ide.signature = r["signature"].is_object().then(|| Signature {
        label: r["signature"]["label"].as_str().unwrap_or("").to_string(),
        params: r["signature"]["params"].as_array().into_iter().flatten().filter_map(Json::as_str).map(str::to_string).collect(),
        active: usize::try_from(r["signature"]["active"].as_u64().unwrap_or(0)).unwrap_or(0),
        doc: r["signature"]["doc"].as_str().unwrap_or("").to_string(),
    });
}

/// Accepts suggestion `i`: the word being typed becomes it, and a snippet's caret lands at `$0`.
fn accept(ide: &mut Ide, text: &mut String, i: usize) {
    let Some(p) = ide.popup.take() else { return };
    let Some(item) = p.items.get(i) else { return };
    let (new, caret) = model::accept(text, p.from, ide.selection.0, &item.insert);
    let before = std::mem::replace(text, new);
    apply_edit(ide, &before, text);
    ide.reveal = Some(caret);
    ide.focus_editor = true;
    ide.signature = None;
}

/// The kind of a suggestion: its letter and colour.
fn kind_badge(kind: &str) -> (&'static str, Cat) {
    match kind {
        "keyword" => ("K", Cat::Purple),
        "snippet" => ("S", Cat::Amber),
        "function" => ("F", Cat::Blue),
        "constant" => ("C", Cat::Teal),
        "variable" => ("V", Cat::Green),
        "property" => ("P", Cat::Orange),
        "method" => ("M", Cat::Blue),
        "class" => ("C", Cat::Red),
        "procedure" => ("S", Cat::Blue),
        _ => ("T", Cat::Slate),
    }
}

/// The suggestion list. Returns the row a click chose.
fn popup_ui(ui: &mut Ui, ide: &mut Ide, t: &Tokens, id: Id) -> Option<usize> {
    let p = ide.popup.clone()?;
    let scroll = ui.data_mut(|d| d.remove_temp::<bool>(id.with("scroll_to_selected"))).unwrap_or(false);
    let mut chosen = None;
    let font = theme::ui_font(12.5);
    Area::new(Id::unique("celdas.vba_ide.suggestions")).order(Order::Foreground).fixed_pos(p.at + vec2(-4.0, 4.0)).show(ui.ctx(), |ui| {
        Frame::new()
            .fill(t.menu_bg)
            .stroke(egui::Stroke::new(1.0, t.separator))
            .corner_radius(theme::CONTROL_RADIUS)
            .shadow(egui::Shadow { offset: [0, 4], blur: 14, spread: 0, color: t.shadow })
            .inner_margin(4)
            .show(ui, |ui| {
                ui.set_width(340.0);
                egui::ScrollArea::vertical().max_height(ROW * POPUP_ROWS).auto_shrink([false, true]).show(ui, |ui| {
                    ui.spacing_mut().item_spacing.y = 0.0;
                    for (i, item) in p.items.iter().enumerate() {
                        let (rect, resp) = ui.allocate_exact_size(vec2(ui.available_width(), ROW), Sense::click());
                        if i == p.selected {
                            ui.painter().rect_filled(rect, 4.0, t.accent_soft);
                            if scroll {
                                ui.scroll_to_rect(rect, None);
                            }
                        } else if resp.hovered() {
                            ui.painter().rect_filled(rect, 4.0, t.hover);
                        }
                        let (letter, cat) = kind_badge(&item.kind);
                        let badge = egui::Rect::from_center_size(rect.left_center() + vec2(14.0, 0.0), vec2(16.0, 16.0));
                        ui.painter().rect_filled(badge, 4.0, t.cat(cat).gamma_multiply(0.22));
                        ui.painter().text(badge.center(), Align2::CENTER_CENTER, letter, theme::ui_font(10.5), t.cat(cat));
                        ui.painter().text(rect.left_center() + vec2(30.0, 0.0), Align2::LEFT_CENTER, &item.label, font.clone(), t.text);
                        if !item.detail.is_empty() {
                            let clip = ui.painter().with_clip_rect(rect.shrink2(vec2(4.0, 0.0)));
                            clip.text(rect.right_center() - vec2(8.0, 0.0), Align2::RIGHT_CENTER, &item.detail, theme::ui_font(11.5), t.text_dim);
                        }
                        if resp.clicked() {
                            chosen = Some(i);
                        }
                    }
                });
                if let Some(item) = p.items.get(p.selected).filter(|i| !i.doc.is_empty()) {
                    ui.add_space(4.0);
                    ui.separator();
                    ui.add(egui::Label::new(RichText::new(&item.doc).font(theme::ui_font(12.0)).color(t.text_dim)).wrap());
                }
            });
    });
    chosen
}

/// The call being written, over the caret's line, with the parameter the caret is in in bold.
fn signature_ui(ui: &mut Ui, ide: &Ide, t: &Tokens, above: Option<egui::Pos2>) {
    let (Some(sig), Some(pos)) = (&ide.signature, above) else { return };
    let mut job = egui::text::LayoutJob::default();
    let plain = egui::TextFormat { font_id: code_editor::font(), color: t.text, ..Default::default() };
    let active = egui::TextFormat {
        font_id: code_editor::font(),
        color: t.accent_text,
        underline: egui::Stroke::new(1.0, t.accent_text),
        ..Default::default()
    };
    let mut at = 0;
    let mut spans: Vec<(usize, usize, bool)> = Vec::new();
    for (i, p) in sig.params.iter().enumerate() {
        if let Some(off) = sig.label.get(at..).and_then(|rest| rest.find(p.as_str())) {
            spans.push((at + off, at + off + p.len(), i == sig.active));
            at += off + p.len();
        }
    }
    let mut cursor = 0;
    for (a, b, on) in spans {
        job.append(sig.label.get(cursor..a).unwrap_or(""), 0.0, plain.clone());
        job.append(sig.label.get(a..b).unwrap_or(""), 0.0, if on { active.clone() } else { plain.clone() });
        cursor = b;
    }
    job.append(sig.label.get(cursor..).unwrap_or(""), 0.0, plain);
    Area::new(Id::unique("celdas.vba_ide.signature"))
        .order(Order::Foreground)
        .interactable(false)
        .pivot(Align2::LEFT_BOTTOM)
        .fixed_pos(pos - vec2(4.0, 6.0))
        .show(ui.ctx(), |ui| {
            Frame::new().fill(t.menu_bg).stroke(egui::Stroke::new(1.0, t.separator)).corner_radius(theme::CONTROL_RADIUS).inner_margin(6).show(
                ui,
                |ui| {
                    ui.set_max_width(520.0);
                    ui.label(job);
                    if !sig.doc.is_empty() {
                        ui.add(egui::Label::new(RichText::new(&sig.doc).font(theme::ui_font(12.0)).color(t.text_dim)).wrap());
                    }
                },
            );
        });
}

/// The other places the name under the caret is written, highlighted.
fn occurrences(app: &mut SheetApp, ide: &mut Ide, text: &str, selection: (usize, usize)) {
    if selection.0 != selection.1 || text.len() > LIVE_LIMIT {
        ide.occurrences.clear();
        ide.occ_key = None;
        return;
    }
    let hash = ide.analysed.as_ref().map_or(0, |(_, h)| *h);
    let key = (ide.current.clone(), hash, selection.0);
    if ide.occ_key.as_ref() == Some(&key) {
        return;
    }
    ide.occ_key = Some(key);
    let (line, col) = model::line_col(text, selection.0);
    let Some((range, _)) = model::word_at(text, selection.0) else {
        ide.occurrences.clear();
        return;
    };
    let _ = range;
    let Ok(r) = app.session.run("vba.occurrences", json!({"code": text, "line": line, "col": col})) else {
        ide.occurrences.clear();
        return;
    };
    let places: Vec<Range<usize>> = r["places"]
        .as_array()
        .into_iter()
        .flatten()
        .filter_map(|p| {
            let line = u32::try_from(p["line"].as_u64()?).ok()?;
            let (a, b) = (usize::try_from(p["start"].as_u64()?).ok()?, usize::try_from(p["end"].as_u64()?).ok()?);
            mark_range(text, line, a, b)
        })
        .collect();
    ide.occurrences = if places.len() > 1 { places } else { Vec::new() };
}

/// Hover help: the mistake under the pointer, the value of a variable of the paused macro, or
/// what the name is.
fn hover(app: &mut SheetApp, ui: &mut Ui, ide: &mut Ide, text: &str, hovered: Option<(usize, egui::Pos2)>, t: &Tokens) {
    if ide.popup.is_some() {
        ide.card = None;
        return;
    }
    let Some((idx, pos)) = hovered else {
        ide.card = None;
        ide.card_key = None;
        return;
    };
    let hash = ide.analysed.as_ref().map_or(0, |(_, h)| *h);
    let key = (ide.current.clone(), hash, idx);
    if ide.card_key.as_ref() != Some(&key) {
        ide.card_key = Some(key);
        ide.card = card_at(app, ide, text, idx, pos);
    }
    let Some(card) = ide.card.clone() else { return };
    Area::new(Id::unique("celdas.vba_ide.hover")).order(Order::Tooltip).interactable(false).fixed_pos(pos + vec2(14.0, 18.0)).show(ui.ctx(), |ui| {
        Frame::new()
            .fill(t.menu_bg)
            .stroke(egui::Stroke::new(1.0, if card.error { t.danger } else { t.separator }))
            .corner_radius(theme::CONTROL_RADIUS)
            .inner_margin(8)
            .show(ui, |ui| {
                ui.set_max_width(460.0);
                let color = if card.error { t.danger } else { t.text };
                ui.add(egui::Label::new(RichText::new(&card.title).font(code_editor::font()).color(color)).wrap());
                if !card.doc.is_empty() {
                    ui.add(egui::Label::new(RichText::new(&card.doc).font(theme::ui_font(12.0)).color(t.text_dim)).wrap());
                }
            });
    });
}

fn card_at(app: &mut SheetApp, ide: &Ide, text: &str, idx: usize, at: egui::Pos2) -> Option<Card> {
    let (line, col) = model::line_col(text, idx);
    if let Some(d) = ide.analysis.diags.iter().find(|d| d.line == line && d.start <= col && col < d.end.max(d.start + 1)) {
        return Some(Card { title: d.message.clone(), doc: String::new(), error: d.severity == Severity::Error, at });
    }
    let (_, word) = model::word_at(text, idx)?;
    if let Some(frame) = ide.debug.as_ref().and_then(|d| d["frames"].get(ide.frame)) {
        let vars = frame["locals"].as_array().into_iter().flatten().chain(ide.debug.as_ref()?["globals"].as_array().into_iter().flatten());
        if let Some(v) = vars.into_iter().find(|v| v["name"].as_str().is_some_and(|n| n.eq_ignore_ascii_case(&word))) {
            let value = v["value"].as_str().unwrap_or("");
            return Some(Card {
                title: format!("{} = {value}", v["name"].as_str().unwrap_or(&word)),
                doc: v["type"].as_str().unwrap_or("").to_string(),
                error: false,
                at,
            });
        }
    }
    let r = app.session.run("vba.hover", json!({"code": text, "line": line, "col": col})).ok()?;
    let title = r["signature"].as_str()?.to_string();
    Some(Card { title, doc: r["doc"].as_str().unwrap_or("").to_string(), error: false, at })
}

/// F12 and Ctrl+click: the caret goes to where the name under it is declared.
pub fn go_to_definition(app: &mut SheetApp, ide: &mut Ide) {
    let name = ide.current.clone();
    let text = ide.text_of(&name);
    let (line, col) = model::line_col(&text, ide.selection.0);
    let Ok(r) = app.session.run("vba.hover", json!({"code": text, "line": line, "col": col})) else { return };
    let d = &r["definition"];
    if let (Some(l), Some(c)) = (d["line"].as_u64().and_then(|l| u32::try_from(l).ok()), d["col"].as_u64().and_then(|c| usize::try_from(c).ok())) {
        ide.go_to(&name, l, c);
    } else {
        ide.status = Some(app.ui.language.tr("No definition found here.").to_string());
    }
}

/// Lays the module out the way the language is written (indentation by block, keywords spelled
/// the usual way). One undo step of the text field is not kept: the change is a draft like typing.
pub fn format(app: &mut SheetApp, ide: &mut Ide) {
    if ide.paused() {
        return;
    }
    let before = ide.text_of(&ide.current.clone());
    let Ok(r) = app.session.run("vba.format", json!({"code": before})) else { return };
    let Some(after) = r["code"].as_str().map(model::normalize) else { return };
    if after != before {
        apply_edit(ide, &before, &after);
        let caret = ide.selection.0.min(after.chars().count());
        ide.reveal = Some(caret);
        ide.status = Some(app.ui.language.tr("Formatted the module.").to_string());
    } else {
        ide.status = Some(app.ui.language.tr("The module is already formatted.").to_string());
    }
    ide.focus_editor = true;
}

/// Moves to the next (or previous) hit of the search from the caret.
pub fn find_step(ide: &mut Ide, forward: bool) {
    let text = ide.text_of(&ide.current.clone());
    let hits = model::find_all(&text, &ide.find.query, ide.find.how);
    if hits.is_empty() {
        return;
    }
    let caret = model::byte_of_char(&text, ide.selection.0);
    let on_hit = hits.iter().any(|r| r.start == caret);
    let next = if forward {
        hits.iter().position(|r| if on_hit { r.start > caret } else { r.start >= caret }).unwrap_or(0)
    } else {
        hits.iter().rposition(|r| r.start < caret).unwrap_or(hits.len() - 1)
    };
    ide.find.index = next;
    if let Some(r) = hits.get(next) {
        ide.reveal = Some(model::char_of_byte(&text, r.start));
    }
}

/// The find and replace bar above the code.
fn find_bar(app: &mut SheetApp, ui: &mut Ui, ide: &mut Ide, text: &mut String) {
    let t = Tokens::get(ui.ctx());
    let lang = app.ui.language;
    let hits = if ide.find.query.is_empty() { Vec::new() } else { model::find_all(text, &ide.find.query, ide.find.how) };
    let mut close = false;
    let mut step = None;
    let mut replace_one = false;
    let mut replace_all = false;
    Frame::new().fill(t.ribbon).inner_margin(egui::Margin::symmetric(10, 6)).show(ui, |ui| {
        ui.horizontal(|ui| {
            let field = ui.add(
                egui::TextEdit::singleline(&mut ide.find.query).hint_text(lang.tr("Find")).desired_width(220.0).id(Id::unique("celdas.vba_ide.find")),
            );
            if ide.find.focus {
                field.request_focus();
                ide.find.focus = false;
            }
            if field.changed() {
                ide.find.index = 0;
            }
            if field.lost_focus() && ui.input(|i| i.key_pressed(Key::Enter)) {
                step = Some(!ui.input(|i| i.modifiers.shift));
                field.request_focus();
            }
            let count = if ide.find.query.is_empty() {
                String::new()
            } else if hits.is_empty() {
                lang.tr("No results").to_string()
            } else {
                lang.tr_fill("{n} of {total}", &[("n", &(ide.find.index.min(hits.len() - 1) + 1).to_string()), ("total", &hits.len().to_string())])
            };
            ui.label(RichText::new(count).font(theme::ui_font(12.0)).color(if !ide.find.query.is_empty() && hits.is_empty() {
                t.danger
            } else {
                t.text_dim
            }));
            if crate::widgets::icon_button(ui, crate::icons::Icon::ChevronUp, t.text, lang.tr("Previous match"), vec2(24.0, 24.0)).clicked() {
                step = Some(false);
            }
            if crate::widgets::icon_button(ui, crate::icons::Icon::Chevron, t.text, lang.tr("Next match"), vec2(24.0, 24.0)).clicked() {
                step = Some(true);
            }
            ui.toggle_value(&mut ide.find.how.case, "Aa").on_hover_text(lang.tr("Match case"));
            ui.toggle_value(&mut ide.find.how.whole, lang.tr("Word")).on_hover_text(lang.tr("Whole word"));
            ui.toggle_value(&mut ide.find.replace, lang.tr("Replace")).on_hover_text(lang.tr("Show replace"));
            ui.with_layout(egui::Layout::right_to_left(egui::Align::Center), |ui| {
                if crate::widgets::icon_button(ui, crate::icons::Icon::Close, t.text_dim, lang.tr("Close"), vec2(24.0, 24.0)).clicked() {
                    close = true;
                }
            });
        });
        if ide.find.replace {
            ui.horizontal(|ui| {
                ui.add(egui::TextEdit::singleline(&mut ide.find.with).hint_text(lang.tr("Replace with")).desired_width(220.0));
                let can = !hits.is_empty() && !ide.paused();
                if ui.add_enabled(can, egui::Button::new(lang.tr("Replace"))).clicked() {
                    replace_one = true;
                }
                if ui.add_enabled(can, egui::Button::new(lang.tr("Replace All"))).clicked() {
                    replace_all = true;
                }
            });
        }
    });
    if close {
        ide.find.open = false;
        ide.focus_editor = true;
    }
    if let Some(forward) = step {
        find_step(ide, forward);
    }
    if replace_one && let Some(r) = hits.get(ide.find.index.min(hits.len().saturating_sub(1))) {
        let before = text.clone();
        text.replace_range(r.clone(), &ide.find.with);
        apply_edit(ide, &before, text);
    }
    if replace_all {
        let before = text.clone();
        let (after, n) = model::replace_all(text, &ide.find.query, &ide.find.with, ide.find.how);
        *text = after;
        apply_edit(ide, &before, text);
        ide.status = Some(lang.tr_fill("Replaced {n} occurrences.", &[("n", &n.to_string())]));
    }
}
