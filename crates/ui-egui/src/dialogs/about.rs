use egui::TextureHandle;
use nodalia_theme::about as ndl;
use nodalia_theme::creator::Look;

use super::Dialog;
use crate::theme::{self, Tokens};

/// Celdas's icon for the About dialog, decoded once.
fn icon(ctx: &egui::Context) -> Option<TextureHandle> {
    let id = egui::Id::NULL.with("celdas.about.icon");
    if let Some(t) = ctx.data(|d| d.get_temp::<TextureHandle>(id)) {
        return Some(t);
    }
    let bytes = include_bytes!("../../../../assets/app-icon/hicolor/128x128/apps/dev.danielmiguelt.Celdas.png");
    let rgba = image::load_from_memory(bytes).ok()?.to_rgba8();
    let size = [rgba.width() as usize, rgba.height() as usize];
    let image = egui::ColorImage::from_rgba_unmultiplied(size, rgba.as_raw());
    let texture = ctx.load_texture("celdas-about-icon", image, egui::TextureOptions::LINEAR);
    ctx.data_mut(|d| d.insert_temp(id, texture.clone()));
    Some(texture)
}

/// Help ▸ About Celdas, as every Nodalia app shows it (`nodalia_theme::about`): the page with the
/// icon, the name and version, the suite and the links. Work by other authors is credited in NOTICE.
pub(super) fn about(ui: &mut egui::Ui, _d: &mut Dialog) {
    let t = Tokens::get(ui.ctx());
    let lang = crate::i18n::current(ui.ctx());
    let look = Look::new(theme::palette_of(&t)).strong_family(theme::ui_bold(1.0).family).medium_family(theme::ui_font(1.0).family);
    about_page(ui, &look, &lang);
}

fn about_page(ui: &mut egui::Ui, look: &Look, lang: &crate::i18n::Language) {
    let version = format!("{} {}", lang.tr("Version"), env!("CARGO_PKG_VERSION"));
    let links = [
        ndl::Link { label: lang.tr("Website"), url: "https://getnodalia.com/celdas" },
        ndl::Link { label: lang.tr("Licence: MIT or Apache-2.0"), url: "https://github.com/getnodalia/celdas/blob/main/LICENSE-MIT" },
    ];
    let info = ndl::Info {
        name: "Celdas",
        version: &version,
        tagline: lang.tr("The spreadsheet of Nodalia's office apps."),
        suite: lang.tr("Part of Nodalia"),
        links: &links,
        footer: lang.tr("Includes work by the ArtCraft Team and other open-source authors (see NOTICE)."),
    };
    let icon = icon(ui.ctx());
    let page = ndl::page(ui, look, &info, |ui, rect| {
        if let Some(tex) = &icon {
            egui::Image::new(tex).fit_to_exact_size(rect.size()).paint_at(ui, rect);
        }
    });
    if let Some(i) = page.picked
        && let Some(link) = links.get(i)
    {
        ui.ctx().open_url(egui::OpenUrl::new_tab(link.url));
    }
}
