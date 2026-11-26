//! Build script: embeds the optional nodalia-fonts build input as `NODALIA_FONTS`.

use std::fmt::Write as _;

#[path = "src/nodalia_fonts_policy.rs"]
mod nodalia_fonts_policy;

/// Build script: emits `nodalia_fonts.rs`; required mode reports a missing or malformed input as a build error.
fn main() {
    emit_nodalia_fonts();
}

/// The optional nodalia-fonts build input (<https://github.com/getnodalia/foundation/tree/nodalia/fonts>, recipe from
/// its `docs/integration.md`). With `NODALIA_FONTS_DIR=<absolute checkout>` the permitted fonts in
/// its `fonts/manifest.txt` are embedded as `NODALIA_FONTS` for `theme`'s fallback chains; unset,
/// the table is empty. Release and acceptance builds also set `NODALIA_FONTS_REQUIRED=1` (exactly
/// `1`; any other value leaves required mode off), which turns a missing checkout, a missing
/// required face or a malformed row into a build error.
///
/// Celdas never copies font bytes into this repository and never reads a fixture from Cargo
/// configuration: the only input is the environment variable. Barred Adobe/Source/Noto CJK faces
/// stay in the shared repository for other consumers but are filtered here. A web build embeds
/// only the explicit `WEB_ALLOWED` faces (no CJK collection). Native builds additionally opt into
/// nodalia-fonts' `fonts/simplified-chinese-manifest.txt` (the OFL LXGW face Lámina uses): it is
/// optional for ordinary builds and required for `NODALIA_FONTS_REQUIRED=1` native builds.
fn emit_nodalia_fonts() {
    println!("cargo::rerun-if-env-changed=NODALIA_FONTS_DIR");
    println!("cargo::rerun-if-env-changed=NODALIA_FONTS_REQUIRED");
    let required = nodalia_fonts_policy::required_mode(std::env::var("NODALIA_FONTS_REQUIRED").ok().as_deref());
    let wasm = std::env::var("CARGO_CFG_TARGET_ARCH").is_ok_and(|a| a == "wasm32");
    let mut src = String::new();
    let _ = writeln!(src, "pub static NODALIA_FONTS_REQUIRED: bool = {required};");
    src.push_str("pub static NODALIA_FONTS: &[NodaliaFont] = &[\n");
    match std::env::var_os("NODALIA_FONTS_DIR").map(std::path::PathBuf::from) {
        Some(dir) => match nodalia_fonts(&dir, wasm, required) {
            Ok(entries) => src.push_str(&entries),
            Err(e) if required => {
                println!("cargo::error=NODALIA_FONTS_DIR={}: {e}", dir.display());
            }
            Err(e) => println!("cargo::warning=building without nodalia-fonts: NODALIA_FONTS_DIR={}: {e}", dir.display()),
        },
        None if required => {
            println!("cargo::error=NODALIA_FONTS_REQUIRED=1 requires NODALIA_FONTS_DIR (an absolute nodalia-fonts checkout)");
        }
        None => {}
    }
    src.push_str("];\n");
    let out = std::path::PathBuf::from(std::env::var_os("OUT_DIR").unwrap_or_default()).join("nodalia_fonts.rs");
    if let Err(e) = std::fs::write(&out, src) {
        println!("cargo::error=writing {}: {e}", out.display());
    }
}

/// One `NodaliaFont { .. }` initialiser per manifest row Celdas may embed.
fn nodalia_fonts(dir: &std::path::Path, wasm: bool, required: bool) -> Result<String, String> {
    // Every file embedded comes from inside the checkout: a manifest row (or a symlink) naming a
    // path outside it would put that file of the build machine into the binary.
    let root = dir.canonicalize().map_err(|e| format!("{}: {e}", dir.display()))?;
    let inside = |path: std::path::PathBuf| -> Result<std::path::PathBuf, String> {
        let real = path.canonicalize().map_err(|e| format!("{}: {e}", path.display()))?;
        if real.starts_with(&root) { Ok(real) } else { Err(format!("{} is outside the nodalia-fonts checkout {}", path.display(), root.display())) }
    };
    // Registered before validation, by the path as given (and its folder, to see the manifest appear
    // or a symlink replaced): a build that skipped the fonts must rebuild once they are there.
    let manifest_path = dir.join("fonts/manifest.txt");
    println!("cargo::rerun-if-changed={}", manifest_path.display());
    if let Some(parent) = manifest_path.parent() {
        println!("cargo::rerun-if-changed={}", parent.display());
    }
    let manifest = inside(manifest_path)?;
    println!("cargo::rerun-if-changed={}", manifest.display());
    let text = std::fs::read_to_string(&manifest).map_err(|e| format!("{}: {e}", manifest.display()))?;
    let rows = nodalia_fonts_policy::parse_manifest(&text)?;
    nodalia_fonts_policy::check_declared_scripts(&rows)?;
    if required {
        nodalia_fonts_policy::check_required(&rows)?;
    }
    let chinese = if wasm { Vec::new() } else { chinese_font_rows(&root, required)? };
    for row in rows.iter().chain(&chinese) {
        if nodalia_fonts_policy::barred(&row.family) {
            println!("cargo::warning=nodalia-fonts: not embedding {} {} (repository asset policy)", row.family, row.style);
        }
    }
    let mut out = String::new();
    for row in nodalia_fonts_policy::embeddable(&rows, wasm).into_iter().chain(nodalia_fonts_policy::embeddable(&chinese, wasm)) {
        let path = inside(root.join(&row.file)).map_err(|e| format!("{}: {e}", row.file))?;
        println!("cargo::rerun-if-changed={}", path.display());
        let scripts: Vec<String> = row.scripts.iter().map(|s| format!("{s:?}")).collect();
        let family = &row.family;
        let style = &row.style;
        let _ = writeln!(
            out,
            "    NodaliaFont {{ family: {family:?}, style: {style:?}, scripts: &[{}], bytes: include_bytes!({path:?}) }},",
            scripts.join(", "),
        );
    }
    Ok(out)
}

/// The opt-in native Simplified Chinese rows. Older nodalia-fonts checkouts may lack the manifest, so
/// an ordinary build warns and carries on; a required acceptance build fails rather than passing
/// Chinese coverage only because the face was silently absent.
fn chinese_font_rows(dir: &std::path::Path, required: bool) -> Result<Vec<nodalia_fonts_policy::ManifestRow>, String> {
    let manifest = dir.join("fonts/simplified-chinese-manifest.txt");
    println!("cargo::rerun-if-changed={}", manifest.display());
    let optional = |e: String| -> Result<Vec<nodalia_fonts_policy::ManifestRow>, String> {
        if required {
            Err(format!("{}: {e}", manifest.display()))
        } else {
            println!("cargo::warning=skipping optional Chinese fonts: {}: {e}", manifest.display());
            Ok(Vec::new())
        }
    };
    // `dir` is the canonical checkout: the manifest itself must not lead out of it either.
    if manifest.canonicalize().is_ok_and(|real| !real.starts_with(dir)) {
        return optional(format!("{} is outside the nodalia-fonts checkout", manifest.display()));
    }
    match std::fs::read_to_string(&manifest) {
        Ok(text) => match nodalia_fonts_policy::parse_manifest(&text).and_then(|rows| {
            nodalia_fonts_policy::check_declared_scripts(&rows)?;
            nodalia_fonts_policy::check_chinese_required(&rows)?;
            Ok(rows)
        }) {
            Ok(rows) => Ok(rows),
            Err(e) => optional(e),
        },
        Err(e) => optional(e.to_string()),
    }
}
