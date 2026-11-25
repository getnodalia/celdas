//! Writing a project's modules back into the `vbaProject.bin` they came from.
//!
//! Everything the file holds besides the modules' source stays as it was: the project's other
//! streams, a form's storages, the protection state. What changes is the modules the caller names,
//! the `dir` stream that lists them and the two streams ([`PROJECT`](rewrite_bin) and
//! `PROJECTwm`) that repeat their names. The compiled form of the code (the performance caches:
//! the bytes before each module's source, the `__SRP_*` streams and the body of `_VBA_PROJECT`) is
//! dropped, as it would describe the old source; the application that opens the file compiles the
//! source again, which is what [MS-OVBA] tells a reader to do when the caches are absent.

use std::collections::HashMap;
use std::ops::Range;

use nodalia_cfb::{CompoundFile, EntryKind, NodeId, Version, Writer};

use crate::project::{MAX_STREAM, ModuleKind, SourceModule, compress_literals, decode_text, decompress, read_modules, u16_at, u32_at, utf16};

/// `_VBA_PROJECT` with no performance cache: the reserved words, version 0xFFFF ("none") and the
/// two reserved fields.
const EMPTY_VBA_PROJECT: [u8; 7] = [0xCC, 0x61, 0xFF, 0xFF, 0x00, 0x00, 0x00];

struct Rec {
    id: u16,
    start: usize,
    data: Range<usize>,
    end: usize,
}

fn records(dir: &[u8]) -> Result<Vec<Rec>, String> {
    let mut out = Vec::new();
    let mut pos = 0;
    while pos < dir.len() {
        let (Some(id), Some(size)) = (u16_at(dir, pos), u32_at(dir, pos + 2)) else {
            return Err("the VBA project's directory is cut short".into());
        };
        let data_start = pos + 6;
        let data_end = data_start.checked_add(size as usize).filter(|e| *e <= dir.len()).ok_or("the VBA project's directory is damaged")?;
        // PROJECTVERSION carries two more bytes than its size says.
        let end = if id == 0x09 { (data_end + 2).min(dir.len()) } else { data_end };
        out.push(Rec { id, start: pos, data: data_start..data_end, end });
        pos = end;
    }
    Ok(out)
}

/// One module's records in `dir`: from MODULENAME to the terminator.
struct Group {
    name: String,
    stream: String,
    first: usize,
    last: usize,
}

fn groups(dir: &[u8], recs: &[Rec], code_page: u16) -> Vec<Group> {
    let mut out: Vec<Group> = Vec::new();
    let mut cur: Option<Group> = None;
    for (i, r) in recs.iter().enumerate() {
        let data = dir.get(r.data.clone()).unwrap_or_default();
        match r.id {
            0x19 => cur = Some(Group { name: decode_text(code_page, data), stream: String::new(), first: i, last: i }),
            0x47 => {
                if let Some(g) = cur.as_mut() {
                    g.name = utf16(data);
                }
            }
            0x1A => {
                if let Some(g) = cur.as_mut() {
                    g.stream = decode_text(code_page, data);
                }
            }
            0x32 => {
                if let Some(g) = cur.as_mut() {
                    g.stream = utf16(data);
                }
            }
            0x2B => {
                if let Some(mut g) = cur.take() {
                    g.last = i;
                    out.push(g);
                }
            }
            _ => {}
        }
    }
    out
}

/// The bytes of `s` in the single-byte `code_page`, or the first character it has no byte for.
pub(crate) fn encode_text(code_page: u16, s: &str) -> Result<Vec<u8>, char> {
    let mut table: HashMap<char, u8> = HashMap::new();
    for b in (0..=255u8).rev() {
        table.insert(nodalia_codepage::code_page_byte(code_page, b), b);
    }
    s.chars().map(|c| if c.is_ascii() { Ok(c as u8) } else { table.get(&c).copied().ok_or(c) }).collect()
}

fn record(out: &mut Vec<u8>, id: u16, data: &[u8]) {
    out.extend_from_slice(&id.to_le_bytes());
    out.extend_from_slice(&(data.len() as u32).to_le_bytes());
    out.extend_from_slice(data);
}

fn utf16_bytes(s: &str) -> Vec<u8> {
    s.encode_utf16().flat_map(u16::to_le_bytes).collect()
}

/// The records of a module the project did not have.
fn new_module_records(out: &mut Vec<u8>, name: &str, name_bytes: &[u8], kind: ModuleKind) {
    record(out, 0x19, name_bytes);
    record(out, 0x47, &utf16_bytes(name));
    record(out, 0x1A, name_bytes);
    record(out, 0x32, &utf16_bytes(name));
    record(out, 0x1C, &[]);
    record(out, 0x48, &[]);
    record(out, 0x31, &0u32.to_le_bytes());
    record(out, 0x1E, &0u32.to_le_bytes());
    record(out, 0x2C, &0xFFFFu16.to_le_bytes());
    record(out, if kind == ModuleKind::Class { 0x22 } else { 0x21 }, &[]);
    record(out, 0x2B, &[]);
}

/// `source` with every line ending written as `eol` (a lone CR, LF or a CR LF pair each count as one).
fn with_line_endings(source: &str, eol: &str) -> String {
    let mut out = String::with_capacity(source.len() + source.len() / 16);
    let mut chars = source.chars().peekable();
    while let Some(c) = chars.next() {
        match c {
            '\r' => {
                if chars.peek() == Some(&'\n') {
                    chars.next();
                }
                out.push_str(eol);
            }
            '\n' => out.push_str(eol),
            _ => out.push(c),
        }
    }
    out
}

/// Which lines of the `PROJECT` stream name a module the project has (`Module=`, `Class=`).
fn declares_module(line: &[u8], name: &[u8]) -> bool {
    [&b"Module="[..], &b"Class="[..]].iter().any(|p| line.strip_prefix(*p).is_some_and(|rest| rest.eq_ignore_ascii_case(name)))
}

fn is_declaration(line: &[u8]) -> bool {
    [&b"Module="[..], &b"Class="[..], &b"Document="[..], &b"BaseClass="[..], &b"Package="[..]].iter().any(|p| line.starts_with(p))
}

fn patch_project_text(text: &[u8], removed: &[Vec<u8>], added: &[(Vec<u8>, ModuleKind)]) -> Vec<u8> {
    let mut lines: Vec<Vec<u8>> = text.split(|b| *b == b'\n').map(|l| l.strip_suffix(b"\r").unwrap_or(l).to_vec()).collect();
    lines.retain(|l| {
        !removed.iter().any(|n| declares_module(l, n)) && !removed.iter().any(|n| l.strip_prefix(n.as_slice()).is_some_and(|r| r.starts_with(b"=")))
    });
    let at = lines
        .iter()
        .rposition(|l| is_declaration(l))
        .map(|i| i + 1)
        .or_else(|| lines.iter().position(|l| l.starts_with(b"ID=")).map(|i| i + 1))
        .unwrap_or(0);
    let new: Vec<Vec<u8>> = added
        .iter()
        .map(|(n, k)| {
            let mut l = if *k == ModuleKind::Class { b"Class=".to_vec() } else { b"Module=".to_vec() };
            l.extend_from_slice(n);
            l
        })
        .collect();
    lines.splice(at..at, new);
    let mut out = Vec::new();
    for (i, l) in lines.iter().enumerate() {
        out.extend_from_slice(l);
        if i + 1 < lines.len() {
            out.extend_from_slice(b"\r\n");
        }
    }
    out
}

/// `PROJECTwm`: each module's name in the project's code page and in UTF-16, then a closing zero.
fn project_wm(names: &[(Vec<u8>, String)]) -> Vec<u8> {
    let mut out = Vec::new();
    for (mbcs, uni) in names {
        out.extend_from_slice(mbcs);
        out.push(0);
        out.extend_from_slice(&utf16_bytes(uni));
        out.extend_from_slice(&[0, 0]);
    }
    out.extend_from_slice(&[0, 0]);
    out
}

/// The `vbaProject.bin` that `bin` becomes when its modules are `modules`: a module of the same
/// name (any case) keeps its place and takes the new source, one the project did not have is
/// added, and one missing from `modules` is removed. A module keeps the kind it had.
///
/// `source` is stored whole, so a module's leading `Attribute` lines (`Attribute VB_Name = …`) are
/// part of it. Lines end with CR LF in the file; a bare LF is written as CR LF.
pub fn rewrite_bin(bin: &[u8], modules: &[SourceModule]) -> Result<Vec<u8>, String> {
    let cf = CompoundFile::open(bin).map_err(|e| format!("the VBA project is not readable: {e}"))?;
    let dir_raw = cf.read_path_capped("VBA/dir", MAX_STREAM as u64).map_err(|e| format!("the VBA project has no directory: {e}"))?;
    let dir = decompress(&dir_raw).ok_or("the VBA project's directory is damaged")?;
    let recs = records(&dir)?;
    let code_page = recs.iter().find(|r| r.id == 0x03).and_then(|r| u16_at(dir.get(r.data.clone())?, 0)).unwrap_or(1252);
    if nodalia_codepage::is_multibyte_code_page(code_page) || !nodalia_codepage::is_byte_code_page(code_page) {
        return Err(format!("the project's code page ({code_page}) cannot be edited yet"));
    }
    let old = groups(&dir, &recs, code_page);

    let key = |n: &str| n.to_lowercase();
    let mut seen: Vec<String> = Vec::new();
    for m in modules {
        if m.name.is_empty() || !m.name.chars().all(|c| c.is_alphanumeric() || c == '_') || m.name.chars().next().is_some_and(|c| c.is_ascii_digit())
        {
            return Err(format!("{:?} is not a valid module name: use letters, digits and underscores, starting with a letter", m.name));
        }
        if seen.contains(&key(&m.name)) {
            return Err(format!("two modules are named {}", m.name));
        }
        seen.push(key(&m.name));
    }

    // The new directory: modules that stay keep their records with the offset of their source
    // zeroed (there is no cache before it any more), removed ones lose theirs, added ones go at the
    // end of the list.
    let kept: Vec<Option<&SourceModule>> = old.iter().map(|g| modules.iter().find(|m| key(&m.name) == key(&g.name))).collect();
    let added: Vec<&SourceModule> = modules.iter().filter(|m| !old.iter().any(|g| key(&g.name) == key(&m.name))).collect();
    let mut new_dir = Vec::with_capacity(dir.len() + 256 * added.len());
    let count = (old.len() - kept.iter().filter(|k| k.is_none()).count() + added.len()) as u16;
    for (i, r) in recs.iter().enumerate() {
        let bytes = dir.get(r.start..r.end).unwrap_or_default();
        if let Some(gi) = old.iter().position(|g| (g.first..=g.last).contains(&i)) {
            if kept.get(gi).is_some_and(|k| k.is_none()) {
                continue;
            }
            if r.id == 0x31 {
                record(&mut new_dir, 0x31, &0u32.to_le_bytes());
                continue;
            }
        } else if r.id == 0x0F {
            record(&mut new_dir, 0x0F, &count.to_le_bytes());
            continue;
        } else if r.id == 0x10 {
            for m in &added {
                let name_bytes = encode_text(code_page, &m.name).map_err(|c| format!("{c:?} is not in the project's code page"))?;
                new_module_records(&mut new_dir, &m.name, &name_bytes, m.kind);
            }
        }
        new_dir.extend_from_slice(bytes);
    }

    let mut w = Writer::new(Version::V3);
    let mut nodes: HashMap<usize, NodeId> = HashMap::new();
    nodes.insert(0, Writer::ROOT);
    let replaced: Vec<String> = old.iter().map(|g| format!("VBA/{}", g.stream).to_lowercase()).collect();
    let removed_names: Vec<Vec<u8>> =
        old.iter().zip(&kept).filter(|(_, k)| k.is_none()).map(|(g, _)| encode_text(code_page, &g.name).unwrap_or_default()).collect();
    let added_names: Vec<(Vec<u8>, ModuleKind)> = added
        .iter()
        .map(|m| Ok((encode_text(code_page, &m.name).map_err(|c| format!("{c:?} is not in the project's code page"))?, m.kind)))
        .collect::<Result<_, String>>()?;
    for id in cf.walk() {
        let (Some(e), Some(path)) = (cf.entry(id), cf.path(id)) else { continue };
        let lower = path.to_lowercase();
        let parent = e.parent.and_then(|p| nodes.get(&p).copied()).unwrap_or(Writer::ROOT);
        match e.kind {
            EntryKind::Storage => {
                let n = w.storage(parent, &e.name, e.clsid).map_err(|x| x.to_string())?;
                nodes.insert(id, n);
            }
            EntryKind::Stream => {
                let data = if lower == "vba/dir" {
                    compress_literals(&new_dir)
                } else if lower == "vba/_vba_project" {
                    EMPTY_VBA_PROJECT.to_vec()
                } else if lower.starts_with("vba/__srp_") || replaced.contains(&lower) {
                    continue;
                } else if lower == "project" {
                    let text = cf.read_capped(id, MAX_STREAM as u64).map_err(|x| x.to_string())?;
                    patch_project_text(&text, &removed_names, &added_names)
                } else if lower == "projectwm" {
                    let mut names: Vec<(Vec<u8>, String)> = Vec::new();
                    for (g, k) in old.iter().zip(&kept) {
                        if k.is_some() {
                            names.push((
                                encode_text(code_page, &g.name).map_err(|c| format!("{c:?} is not in the project's code page"))?,
                                g.name.clone(),
                            ));
                        }
                    }
                    for m in &added {
                        names
                            .push((encode_text(code_page, &m.name).map_err(|c| format!("{c:?} is not in the project's code page"))?, m.name.clone()));
                    }
                    project_wm(&names)
                } else {
                    cf.read_capped(id, MAX_STREAM as u64).map_err(|x| x.to_string())?
                };
                w.stream(parent, &e.name, data).map_err(|x| x.to_string())?;
            }
            _ => {}
        }
    }
    let olds: HashMap<String, String> = read_modules(bin)?.into_iter().map(|m| (key(&m.name), m.source)).collect();
    // The modules' streams: each source as text in the code page, compressed, with no cache before it.
    for m in modules {
        let stream = old.iter().find(|g| key(&g.name) == key(&m.name)).map_or(m.name.as_str(), |g| g.stream.as_str());
        // A module keeps the line ending its file used (Excel for the Mac writes a bare LF).
        let eol = if olds.get(&key(&m.name)).is_some_and(|o| !o.contains("\r\n") && o.contains('\n')) { "\n" } else { "\r\n" };
        let mut text = with_line_endings(&m.source, eol);
        if !text.is_empty() && !text.ends_with(eol) {
            text.push_str(eol);
        }
        let bytes = encode_text(code_page, &text).map_err(|c| {
            let line = text.split(eol).position(|l| l.contains(c)).map_or(0, |i| i + 1);
            format!("module {}, line {line}: {c:?} is not in the project's code page; use a character the page has, or write it with ChrW", m.name)
        })?;
        let vba = nodes.get(&cf.find("VBA").map_err(|e| e.to_string())?).copied().ok_or("the VBA storage is missing")?;
        w.stream(vba, stream, compress_literals(&bytes)).map_err(|e| e.to_string())?;
    }
    Ok(w.finish())
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::project::{Project, write_bin};

    fn module(name: &str, kind: ModuleKind, source: &str) -> SourceModule {
        SourceModule { name: name.into(), kind, source: source.into() }
    }

    /// A project as Excel keeps one: modules, the `PROJECT` text, a form's storage and a stale cache.
    fn project_with_extras() -> Vec<u8> {
        let base = write_bin(&[
            module("Module1", ModuleKind::Standard, "Sub A()\r\nEnd Sub\r\n"),
            module("Sheet1", ModuleKind::Class, "Option Explicit\r\n"),
        ])
        .unwrap();
        let cf = CompoundFile::open(&base).unwrap();
        let mut w = Writer::new(Version::V3);
        for id in cf.walk() {
            let (Some(e), Some(p)) = (cf.entry(id), cf.path(id)) else { continue };
            if e.kind == EntryKind::Stream {
                w.stream_at(&p, cf.read(id).unwrap()).unwrap();
            }
        }
        w.stream_at("PROJECT", b"ID=\"{00000000-0000-0000-0000-000000000000}\"\r\nModule=Module1\r\nDocument=Sheet1/&H00000000\r\nName=\"VBAProject\"\r\n\r\n[Workspace]\r\nModule1=0, 0, 0, 0, C\r\n".to_vec()).unwrap();
        w.stream_at("VBA/_VBA_PROJECT", vec![0xCC, 0x61, 0xDF, 0, 0, 0x11, 0, 0xFF, 1, 2, 3]).unwrap();
        w.stream_at("VBA/__SRP_0", vec![9; 20]).unwrap();
        w.stream_at("UserForm1/\u{3}VBFrame", b"form".to_vec()).unwrap();
        w.finish()
    }

    fn text_of(bin: &[u8], path: &str) -> Vec<u8> {
        CompoundFile::open(bin).unwrap().read_path(path).unwrap()
    }

    #[test]
    fn an_edited_module_reads_back_with_its_new_source() {
        let bin = project_with_extras();
        let edited = rewrite_bin(
            &bin,
            &[module("Module1", ModuleKind::Standard, "Sub B()\nEnd Sub"), module("Sheet1", ModuleKind::Class, "Option Explicit\r\n")],
        )
        .unwrap();
        let p = Project::from_bin(&edited).unwrap();
        assert_eq!(p.modules.len(), 2);
        assert_eq!(p.modules[0].source, "Sub B()\r\nEnd Sub\r\n");
        assert_eq!(p.procedures()[0].name, "B");
        assert_eq!(p.modules[1].kind, ModuleKind::Class);
    }

    #[test]
    fn everything_else_in_the_file_is_kept_and_the_caches_go() {
        let bin = project_with_extras();
        let edited =
            rewrite_bin(&bin, &[module("Module1", ModuleKind::Standard, "Sub A()\r\nEnd Sub\r\n"), module("Sheet1", ModuleKind::Class, "")]).unwrap();
        assert_eq!(text_of(&edited, "UserForm1/\u{3}VBFrame"), b"form");
        let cf = CompoundFile::open(&edited).unwrap();
        assert!(!cf.has_stream("VBA/__SRP_0"));
        assert_eq!(text_of(&edited, "VBA/_VBA_PROJECT"), EMPTY_VBA_PROJECT);
        assert!(String::from_utf8(text_of(&edited, "PROJECT")).unwrap().contains("Module=Module1\r\nDocument=Sheet1/&H00000000"));
    }

    #[test]
    fn a_module_can_be_added_and_removed() {
        let bin = project_with_extras();
        let added = rewrite_bin(
            &bin,
            &[
                module("Module1", ModuleKind::Standard, "Sub A()\r\nEnd Sub\r\n"),
                module("Sheet1", ModuleKind::Class, ""),
                module("Calc", ModuleKind::Standard, "Function F()\r\n F = 1\r\nEnd Function\r\n"),
            ],
        )
        .unwrap();
        let p = Project::from_bin(&added).unwrap();
        assert_eq!(p.modules.iter().map(|m| m.name.as_str()).collect::<Vec<_>>(), ["Module1", "Sheet1", "Calc"]);
        assert!(p.procedures().iter().any(|x| x.name == "F" && x.is_function));
        assert!(String::from_utf8(text_of(&added, "PROJECT")).unwrap().contains("Module=Calc"));
        let removed = rewrite_bin(
            &added,
            &[module("Sheet1", ModuleKind::Class, ""), module("Calc", ModuleKind::Standard, "Function F()\r\n F = 1\r\nEnd Function\r\n")],
        )
        .unwrap();
        let p = Project::from_bin(&removed).unwrap();
        assert_eq!(p.modules.iter().map(|m| m.name.as_str()).collect::<Vec<_>>(), ["Sheet1", "Calc"]);
        let project_text = String::from_utf8(text_of(&removed, "PROJECT")).unwrap();
        assert!(!project_text.contains("Module=Module1") && project_text.contains("Module=Calc"));
        assert!(!CompoundFile::open(&removed).unwrap().has_stream("VBA/Module1"));
    }

    #[test]
    fn text_outside_the_code_page_is_refused_with_its_line() {
        let bin = project_with_extras();
        let e = rewrite_bin(&bin, &[module("Module1", ModuleKind::Standard, "Sub A()\r\n ' \u{4e2d}\r\nEnd Sub\r\n")]).unwrap_err();
        assert!(e.contains("module Module1, line 2"), "{e}");
        let ok = rewrite_bin(&bin, &[module("Module1", ModuleKind::Standard, "' año\r\n")]).unwrap();
        assert_eq!(Project::from_bin(&ok).unwrap().modules[0].source, "' año\r\n");
    }

    #[test]
    fn bad_module_names_and_hostile_files_are_errors() {
        let bin = project_with_extras();
        assert!(rewrite_bin(&bin, &[module("1bad", ModuleKind::Standard, "")]).is_err());
        assert!(rewrite_bin(&bin, &[module("Two", ModuleKind::Standard, ""), module("TWO", ModuleKind::Standard, "")]).is_err());
        assert!(rewrite_bin(b"not a compound file", &[]).is_err());
    }

    /// With `CELDAS_TURBO_BIN` naming a real `vbaProject.bin`: rewriting it with one line added keeps
    /// every module's text, and `CELDAS_TURBO_OUT` (when set) receives the result to open elsewhere.
    #[test]
    fn a_real_project_survives_an_edit() {
        let Ok(path) = std::env::var("CELDAS_TURBO_BIN") else { return };
        let bin = std::fs::read(path).unwrap();
        let before = Project::from_bin(&bin).unwrap();
        let mut modules: Vec<SourceModule> =
            before.modules.iter().map(|m| SourceModule { name: m.name.clone(), kind: m.kind, source: m.source.clone() }).collect();
        let first = modules.iter_mut().find(|m| m.kind == ModuleKind::Standard).unwrap();
        first.source.push_str("' edited by Celdas\r\n");
        let out = rewrite_bin(&bin, &modules).unwrap();
        let after = Project::from_bin(&out).unwrap();
        assert_eq!(after.modules.len(), before.modules.len());
        for (a, b) in after.modules.iter().zip(&before.modules) {
            assert_eq!(a.name, b.name);
            if a.name != modules.iter().find(|m| m.kind == ModuleKind::Standard).unwrap().name {
                assert_eq!(a.source, b.source, "{}", a.name);
            }
        }
        if let Ok(dest) = std::env::var("CELDAS_TURBO_OUT") {
            std::fs::write(dest, out).unwrap();
        }
    }
}
