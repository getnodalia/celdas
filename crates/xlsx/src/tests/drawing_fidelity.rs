//! Drawing details Excel shows and a reader must keep: DrawingML colour transforms, hidden legend
//! entries and cropped pictures.

use std::sync::Arc;

use celdas_core::CellRef;
use celdas_model::{Anchor, Color, Image, Sheet, Workbook};

use crate::drawing::dml_color;
use crate::xml::parse;
use crate::{read_xlsx, write_xlsx};

fn colour(xml: &str) -> Option<Color> {
    dml_color(&parse(xml.as_bytes()).unwrap())
}

#[test]
fn plain_colours_keep_their_value() {
    assert_eq!(colour(r#"<a:solidFill><a:srgbClr val="70AD47"/></a:solidFill>"#), Some(Color::Rgb(0x70AD47)));
    assert_eq!(colour(r#"<a:solidFill><a:schemeClr val="accent6"/></a:solidFill>"#), Some(Color::Theme(9, 0)));
    assert_eq!(colour(r#"<a:solidFill><a:schemeClr val="bogus"/></a:solidFill>"#), None);
}

#[test]
fn lum_transforms_become_a_tint_on_the_theme_slot() {
    // "Lighter 40%" is lumMod 60% + lumOff 40%.
    let lighter = colour(r#"<a:solidFill><a:schemeClr val="accent6"><a:lumMod val="60000"/><a:lumOff val="40000"/></a:schemeClr></a:solidFill>"#);
    assert_eq!(lighter, Some(Color::Theme(9, 400)));
    // "Darker 25%" is lumMod 75%.
    let darker = colour(r#"<a:solidFill><a:schemeClr val="accent1"><a:lumMod val="75000"/></a:schemeClr></a:solidFill>"#);
    assert_eq!(darker, Some(Color::Theme(4, -250)));
}

#[test]
fn rgb_colours_have_the_transform_applied() {
    let lighter = colour(r#"<a:solidFill><a:srgbClr val="000000"><a:lumMod val="50000"/><a:lumOff val="50000"/></a:srgbClr></a:solidFill>"#);
    assert_eq!(lighter, Some(Color::Rgb(0x808080)));
}

#[test]
fn hostile_transform_values_do_not_panic() {
    for v in ["-99999999999", "NaN", "inf", "abc"] {
        let xml = format!(r#"<a:solidFill><a:schemeClr val="accent2"><a:lumMod val="{v}"/><a:lumOff val="{v}"/></a:schemeClr></a:solidFill>"#);
        let _ = colour(&xml);
    }
}

#[test]
fn picture_crop_survives_a_round_trip() {
    let mut s = Sheet::new("Data");
    let png = b"\x89PNG\r\n\x1a\n0000IHDRfake".to_vec();
    s.images.push(Image {
        id: 1,
        anchor: Anchor { cell: CellRef::new(1, 1), width: 100.0, height: 60.0, ..Anchor::default() },
        data: png.into(),
        mime: "image/png".into(),
        alt: String::new(),
        crop: [0.1, 0.2, 0.25, 0.0],
    });
    let wb = Workbook { sheets: vec![Arc::new(s)], ..Workbook::default() };
    let (back, _) = read_xlsx(&write_xlsx(&wb).unwrap()).unwrap();
    let got = back.sheets[0].images[0].crop;
    for (g, w) in got.iter().zip([0.1, 0.2, 0.25, 0.0]) {
        assert!((g - w).abs() < 1e-3, "{got:?}");
    }
}

#[test]
fn uncropped_pictures_write_no_source_rectangle() {
    let mut s = Sheet::new("Data");
    s.images.push(Image {
        id: 1,
        anchor: Anchor { cell: CellRef::new(1, 1), width: 10.0, height: 10.0, ..Anchor::default() },
        data: b"\x89PNG\r\n\x1a\n0000IHDRfake".to_vec().into(),
        mime: "image/png".into(),
        alt: String::new(),
        crop: [0.0; 4],
    });
    let wb = Workbook { sheets: vec![Arc::new(s)], ..Workbook::default() };
    let bytes = write_xlsx(&wb).unwrap();
    let mut z = zip::ZipArchive::new(std::io::Cursor::new(bytes)).unwrap();
    let mut xml = String::new();
    std::io::Read::read_to_string(&mut z.by_name("xl/drawings/drawing1.xml").unwrap(), &mut xml).unwrap();
    assert!(!xml.contains("srcRect"));
}
