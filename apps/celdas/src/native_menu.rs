//! macOS: the native menu bar (File, Edit, View, Insert, Format, Tools, Data, Window, Help), as
//! a desktop spreadsheet has on the Mac. The menus come from `menu_tree.rs`, and items run the same
//! commands as the ribbon.

use std::str::FromStr;
use std::sync::mpsc::{Receiver, channel};

use celdas_ui_egui::i18n::Language;
use celdas_ui_egui::{OpenWindow, SheetApp};
use muda::accelerator::Accelerator;
use muda::{CheckMenuItem, Menu, MenuEvent, MenuItem, PredefinedMenuItem, Submenu};
use serde_json::json;

use crate::menu_tree::{QUIT, dispatch, label, tree};

pub struct NativeMenu {
    _menu: Menu,
    /// The language the menus were built in; they are built again when the preference changes.
    language: Language,
    events: Receiver<MenuEvent>,
    window_menu: Option<Submenu>,
    /// The Window menu's window items and what they last showed.
    window_items: Vec<CheckMenuItem>,
    windows: Vec<OpenWindow>,
}

/// Menu id of the Window menu item that brings workbook `doc` forward.
const WINDOW_ITEM: &str = "window:id:";

impl NativeMenu {
    pub fn install(ctx: &egui::Context, language: Language) -> NativeMenu {
        let (menu, window_menu) = build(language);
        let (tx, rx) = channel();
        let ctx = ctx.clone();
        MenuEvent::set_event_handler(Some(move |e: MenuEvent| {
            let _ = tx.send(e);
            ctx.request_repaint();
        }));
        NativeMenu { _menu: menu, language, events: rx, window_menu, window_items: Vec::new(), windows: Vec::new() }
    }

    /// Builds the menus again in `language`, and lists the open windows in the new Window menu.
    fn rebuild(&mut self, language: Language) {
        let (menu, window_menu) = build(language);
        self._menu = menu;
        self.language = language;
        self.window_menu = window_menu;
        self.window_items.clear();
        self.windows.clear();
    }

    /// Keeps the Window menu's list of open windows (title, unsaved bullet, active check) in
    /// step with the session.
    fn sync_windows(&mut self, app: &SheetApp) {
        let windows = app.open_windows();
        if windows != self.windows
            && let Some(menu) = &self.window_menu
        {
            let same_windows = windows.len() == self.windows.len() && windows.iter().zip(&self.windows).all(|(a, b)| a.window == b.window);
            if same_windows {
                for (item, w) in self.window_items.iter().zip(&windows) {
                    item.set_text(w.label());
                }
            } else {
                for item in self.window_items.drain(..) {
                    let _ = menu.remove(&item);
                }
                for w in &windows {
                    let item = CheckMenuItem::with_id(format!("{WINDOW_ITEM}{}", w.window), w.label(), true, w.active, None);
                    let _ = menu.append(&item);
                    self.window_items.push(item);
                }
            }
            self.windows = windows;
        }
        // AppKit flips a check item when it is chosen, so the active check is put back every poll.
        for (item, w) in self.window_items.iter().zip(&self.windows) {
            if item.is_checked() != w.active {
                item.set_checked(w.active);
            }
        }
    }

    /// Runs the chosen menu items. Commands go to the active window, the most recently focused;
    /// `viewport` is its egui viewport, for the items that act on its input.
    pub fn poll(&mut self, app: &mut SheetApp, ctx: &egui::Context, viewport: egui::ViewportId) {
        if app.ui.language != self.language {
            self.rebuild(app.ui.language);
        }
        while let Ok(ev) = self.events.try_recv() {
            let id = ev.id.as_ref().to_string();
            if let Some(window) = id.strip_prefix(WINDOW_ITEM).and_then(|w| w.parse::<u64>().ok()) {
                app.run_or_alert("view.switchWindows", json!({"window": window}));
                continue;
            }
            dispatch(app, ctx, viewport, &id);
        }
        self.sync_windows(app);
    }
}

/// The menu bar in `language`, and its Window menu (always found by its English title).
fn build(language: Language) -> (Menu, Option<Submenu>) {
    let menu = Menu::new();
    let app_menu = Submenu::new("Celdas", true);
    let about = MenuItem::with_id("dialog:about", label(language, "About Celdas"), true, None);
    let _ = app_menu.append_items(&[
        &about,
        &PredefinedMenuItem::separator(),
        &MenuItem::with_id("dialog:backstage:options", label(language, "Settings…"), true, Accelerator::from_str("CMD+COMMA").ok()),
        &PredefinedMenuItem::separator(),
        &PredefinedMenuItem::services(None),
        &PredefinedMenuItem::separator(),
        &PredefinedMenuItem::hide(None),
        &PredefinedMenuItem::hide_others(None),
        &PredefinedMenuItem::show_all(None),
        &PredefinedMenuItem::separator(),
        &MenuItem::with_id(QUIT, label(language, "Quit Celdas"), true, Accelerator::from_str("CMD+Q").ok()),
    ]);
    let _ = menu.append(&app_menu);
    let mut window_menu = None;
    for (title, entries) in tree() {
        let sub = Submenu::new(label(language, title), true);
        for (text, id, acc) in entries {
            if text == "-" {
                let _ = sub.append(&PredefinedMenuItem::separator());
                continue;
            }
            let accel = acc.and_then(|a| Accelerator::from_str(a).ok());
            let item = MenuItem::with_id(id, label(language, text), true, accel);
            let _ = sub.append(&item);
        }
        let _ = menu.append(&sub);
        if title == "Window" {
            window_menu = Some(sub);
        }
    }
    menu.init_for_nsapp();
    // NSApp's windowsMenu is where macOS adds Move & Resize, Fill and Center, the targets of the
    // system tiling shortcuts (Fn+Ctrl+arrows, Fn+F). Only once the menu is NSApp's: muda looks
    // the submenu up in the installed main menu.
    if let Some(window) = &window_menu {
        window.set_as_windows_menu_for_nsapp();
    }
    (menu, window_menu)
}
