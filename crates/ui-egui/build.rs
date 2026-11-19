//! Bake the contributor credits into the binary as static Rust tables. Nothing is read at run time.
//!
//! - contributors/contributors.json: commit stats per GitHub username, with a mark for whose
//!   people.toml entry is verified.
//! - contributors/people.toml: the names people chose for themselves. A name is used only when the
//!   JSON says the entry was verified (`people_entry`).
//!
//! A missing or malformed file gives empty tables (or no names) and a warning.

use std::fmt::Write as _;

use serde_json::Value;

#[path = "src/nodalia_fonts_policy.rs"]
mod nodalia_fonts_policy;

/// Build script: reads the contributor JSON/TOML and writes generated `credits.rs`, then emits `nodalia_fonts.rs`; missing or malformed optional inputs warn, while required mode reports build errors.
fn main() {
    // Read when the script runs: `env!` would keep the path from when the script was compiled,
    // which goes stale if the checkout moves (and a missing file reruns the script every build).
    let manifest = std::env::var_os("CARGO_MANIFEST_DIR").map(std::path::PathBuf::from).unwrap_or_default();
    let dir = manifest.join("../../contributors");
    let json = read(&dir.join("contributors.json"), "the About window shows no contributors", |t| {
        serde_json::from_str::<Value>(t).map_err(|e| e.to_string())
    })
    .unwrap_or(Value::Null);
    let people = read(&dir.join("people.toml"), "contributors are shown by username only", |t| t.parse::<toml::Table>().map_err(|e| e.to_string()))
        .unwrap_or_default();
    let out = std::path::PathBuf::from(std::env::var_os("OUT_DIR").unwrap_or_default()).join("credits.rs");
    if let Err(e) = std::fs::write(&out, generate(&json, &people)) {
        println!("cargo::warning=writing {}: {e}", out.display());
    }
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

fn read<T>(path: &std::path::Path, fallback: &str, parse: impl Fn(&str) -> Result<T, String>) -> Option<T> {
    println!("cargo::rerun-if-changed={}", path.display());
    let r = std::fs::read_to_string(path).map_err(|e| e.to_string()).and_then(|t| parse(&t));
    if let Err(e) = &r {
        println!("cargo::warning={}: {e}; {fallback}", path.display());
    }
    r.ok()
}

/// A non-empty string field of someone's people.toml entry, as a Rust `Option<&str>` literal.
fn name(v: Option<&str>) -> String {
    match v.map(str::trim).filter(|s| !s.is_empty()) {
        Some(x) => format!("Some({x:?})"),
        None => "None".into(),
    }
}

fn s(v: &Value, k: &str) -> String {
    format!("{:?}", v.get(k).and_then(Value::as_str).unwrap_or(""))
}

fn n(v: &Value, k: &str) -> u64 {
    v.get(k).and_then(Value::as_u64).unwrap_or(0)
}

fn rows(json: &Value, k: &str) -> Vec<Value> {
    json.get(k).and_then(Value::as_array).cloned().unwrap_or_default()
}

fn generate(json: &Value, people: &toml::Table) -> String {
    let mut o = String::new();
    let _ = writeln!(o, "pub static TOTAL_COMMITS: u64 = {};", n(json, "commits"));
    let _ = writeln!(o, "pub static CONTRIBUTORS: &[Contributor] = &[");
    for c in rows(json, "contributors") {
        let login = c.get("login").and_then(Value::as_str).unwrap_or("");
        if login.is_empty() {
            continue;
        }
        // Names only from a verified, self-chosen people.toml entry.
        let entry = people.get("people").and_then(|p| p.get(login)).filter(|_| c.get("people_entry").and_then(Value::as_bool) == Some(true));
        let field = |k: &str| entry.and_then(|e| e.get(k)).and_then(toml::Value::as_str);
        let synced = entry.and_then(|e| e.get("sync_github_name")).and_then(toml::Value::as_bool) == Some(true);
        let display = field("display_name")
            .filter(|s| !s.trim().is_empty())
            .or_else(|| if synced { c.get("github_name").and_then(Value::as_str) } else { None });
        let _ = writeln!(
            o,
            "    Contributor {{ login: {login:?}, display_name: {}, real_name: {}, prs: {}, commits: {}, lines_added: {}, lines_deleted: {}, binary_added: {}, binary_deleted: {}, first_commit: {}, last_commit: {} }},",
            name(display),
            name(field("real_name")),
            n(&c, "prs"),
            n(&c, "commits"),
            n(&c, "lines_added"),
            n(&c, "lines_deleted"),
            n(&c, "binary_added"),
            n(&c, "binary_deleted"),
            s(&c, "first_commit"),
            s(&c, "last_commit"),
        );
    }
    let _ = writeln!(o, "];");
    let _ = writeln!(o, "pub static MODELS: &[Model] = &[");
    for m in rows(json, "models") {
        let _ = writeln!(
            o,
            "    Model {{ company: {}, model: {}, version: {}, commits: {}, lines_added: {}, lines_deleted: {} }},",
            s(&m, "company"),
            s(&m, "model"),
            s(&m, "version"),
            n(&m, "commits"),
            n(&m, "lines_added"),
            n(&m, "lines_deleted"),
        );
    }
    let _ = writeln!(o, "];");
    o
}
