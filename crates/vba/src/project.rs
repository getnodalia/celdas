//! A workbook's VBA project (`vbaProject.bin`, [MS-OVBA]): its modules' source, and the parsed
//! procedures.

use nodalia_cfb::CompoundFile;

use crate::ast::{Module, Proc, ProcKind};
use crate::lexer::SyntaxError;
use crate::parser::parse_module;

/// Most bytes a decompressed stream may have (a hostile file must not exhaust memory).
pub(crate) const MAX_STREAM: usize = 64 << 20;

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum ModuleKind {
    /// A standard module (`.bas`).
    Standard,
    /// A class module, or the code behind a sheet or the workbook.
    Class,
}

/// A module as stored: its name and source text.
#[derive(Clone, Debug)]
pub struct SourceModule {
    pub name: String,
    pub kind: ModuleKind,
    pub source: String,
}

pub struct LoadedModule {
    pub name: String,
    pub kind: ModuleKind,
    pub source: String,
    /// The parsed module, or why it could not be parsed.
    pub parsed: Result<Module, SyntaxError>,
}

/// One procedure of the project, as listed to a user.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct ProcInfo {
    pub module: String,
    pub name: String,
    pub is_function: bool,
    pub public: bool,
    pub has_params: bool,
}

pub struct Project {
    pub modules: Vec<LoadedModule>,
}

impl Project {
    pub fn from_sources(sources: Vec<SourceModule>) -> Project {
        let modules = sources
            .into_iter()
            .map(|m| {
                let parsed = parse_module(&m.name, &m.source);
                LoadedModule { name: m.name, kind: m.kind, source: m.source, parsed }
            })
            .collect();
        Project { modules }
    }

    /// Reads the project from the bytes of `vbaProject.bin`.
    pub fn from_bin(bin: &[u8]) -> Result<Project, String> {
        Ok(Project::from_sources(read_modules(bin)?))
    }

    /// The public and private procedures of every module that parsed.
    pub fn procedures(&self) -> Vec<ProcInfo> {
        let mut out = Vec::new();
        for m in &self.modules {
            let Ok(parsed) = &m.parsed else { continue };
            for p in &parsed.procs {
                if !matches!(p.kind, ProcKind::Sub | ProcKind::Function) {
                    continue;
                }
                out.push(ProcInfo {
                    module: m.name.clone(),
                    name: p.display.clone(),
                    is_function: p.kind == ProcKind::Function,
                    public: p.public,
                    has_params: !p.params.is_empty(),
                });
            }
        }
        out
    }

    /// Modules that could not be parsed, with the error.
    pub fn syntax_errors(&self) -> Vec<(String, SyntaxError)> {
        self.modules.iter().filter_map(|m| m.parsed.as_ref().err().map(|e| (m.name.clone(), e.clone()))).collect()
    }

    pub fn module_index(&self, name: &str) -> Option<usize> {
        self.modules.iter().position(|m| m.name.eq_ignore_ascii_case(name))
    }

    /// A procedure by name (lower case), looking first in module `prefer`.
    pub fn find_proc(&self, name: &str, prefer: Option<usize>) -> Option<(usize, &Proc)> {
        let find_in = |idx: usize| -> Option<(usize, &Proc)> {
            let m = self.modules.get(idx)?.parsed.as_ref().ok()?;
            m.procs.iter().find(|p| &*p.name == name && matches!(p.kind, ProcKind::Sub | ProcKind::Function)).map(|p| (idx, p))
        };
        if let Some(i) = prefer
            && let Some(hit) = find_in(i)
        {
            return Some(hit);
        }
        (0..self.modules.len()).find_map(find_in)
    }

    /// A property procedure (`Property Get`/`Let`/`Set`).
    pub fn find_property(&self, name: &str, kind: ProcKind) -> Option<(usize, &Proc)> {
        self.modules
            .iter()
            .enumerate()
            .find_map(|(i, m)| m.parsed.as_ref().ok()?.procs.iter().find(|p| &*p.name == name && p.kind == kind).map(|p| (i, p)))
    }
}

// ---- reading vbaProject.bin ----

pub(crate) fn u16_at(b: &[u8], i: usize) -> Option<u16> {
    Some(u16::from_le_bytes([*b.get(i)?, *b.get(i + 1)?]))
}

pub(crate) fn u32_at(b: &[u8], i: usize) -> Option<u32> {
    Some(u32::from_le_bytes([*b.get(i)?, *b.get(i + 1)?, *b.get(i + 2)?, *b.get(i + 3)?]))
}

/// Decompresses an MS-OVBA compressed container.
pub fn decompress(data: &[u8]) -> Option<Vec<u8>> {
    if *data.first()? != 1 {
        return None;
    }
    let mut out: Vec<u8> = Vec::new();
    let mut pos = 1;
    while pos < data.len() {
        let header = u16_at(data, pos)?;
        let chunk_end = (pos + usize::from(header & 0x0FFF) + 3).min(data.len());
        let compressed = header & 0x8000 != 0;
        pos += 2;
        let chunk_start = out.len();
        if !compressed {
            out.extend_from_slice(data.get(pos..(pos + 4096).min(data.len()))?);
            pos += 4096;
        } else {
            while pos < chunk_end {
                let flags = *data.get(pos)?;
                pos += 1;
                for bit in 0..8 {
                    if pos >= chunk_end {
                        break;
                    }
                    if flags >> bit & 1 == 0 {
                        out.push(*data.get(pos)?);
                        pos += 1;
                    } else {
                        let token = u16_at(data, pos)?;
                        pos += 2;
                        copy_token(&mut out, chunk_start, token)?;
                    }
                }
            }
        }
        if out.len() > MAX_STREAM {
            return None;
        }
        pos = pos.max(chunk_end);
    }
    Some(out)
}

fn copy_token(out: &mut Vec<u8>, chunk_start: usize, token: u16) -> Option<()> {
    let done = out.len() - chunk_start;
    let bit_count = if done <= 16 { 4 } else { 32 - (done as u32 - 1).leading_zeros() };
    let length_mask = 0xFFFFu16 >> bit_count;
    let length = usize::from(token & length_mask) + 3;
    let offset = usize::from((token & !length_mask) >> (16 - bit_count)) + 1;
    if offset > done {
        return None;
    }
    for _ in 0..length {
        let b = *out.get(out.len() - offset)?;
        out.push(b);
    }
    Some(())
}

/// Compresses `data` into an MS-OVBA container using literal tokens only: valid, not small. For
/// building test projects and for writing modules back.
pub fn compress_literals(data: &[u8]) -> Vec<u8> {
    let mut out = vec![1u8];
    for chunk in data.chunks(3640) {
        let mut body = Vec::with_capacity(chunk.len() + chunk.len() / 8 + 1);
        for group in chunk.chunks(8) {
            body.push(0);
            body.extend_from_slice(group);
        }
        // Size field = length of the chunk including the header, minus 3; flag bit set.
        let header = 0xB000u16 | ((body.len() + 2 - 3) as u16 & 0x0FFF);
        out.extend_from_slice(&header.to_le_bytes());
        out.extend_from_slice(&body);
    }
    out
}

struct DirModule {
    name: String,
    stream: String,
    offset: u32,
    kind: ModuleKind,
}

pub(crate) fn decode_text(code_page: u16, bytes: &[u8]) -> String {
    bytes.iter().map(|b| nodalia_codepage::code_page_byte(code_page, *b)).collect()
}

pub(crate) fn utf16(bytes: &[u8]) -> String {
    let units: Vec<u16> = bytes.as_chunks::<2>().0.iter().map(|c| u16::from_le_bytes(*c)).collect();
    String::from_utf16_lossy(&units)
}

fn parse_dir(dir: &[u8]) -> Vec<(DirModule, u16)> {
    let mut code_page = 1252u16;
    let mut mods: Vec<(DirModule, u16)> = Vec::new();
    let mut cur: Option<DirModule> = None;
    let mut pos = 0;
    while let (Some(id), Some(size)) = (u16_at(dir, pos), u32_at(dir, pos + 2)) {
        pos += 6;
        let size = size as usize;
        // PROJECTVERSION carries two more bytes than its size says.
        let next = if id == 0x09 { pos + size + 2 } else { pos + size };
        let Some(data) = dir.get(pos..pos + size) else { break };
        match id {
            0x03 => code_page = u16_at(data, 0).unwrap_or(1252),
            0x19 => cur = Some(DirModule { name: decode_text(code_page, data), stream: String::new(), offset: 0, kind: ModuleKind::Standard }),
            0x47 => {
                if let Some(m) = cur.as_mut() {
                    m.name = utf16(data);
                }
            }
            0x1A => {
                if let Some(m) = cur.as_mut() {
                    m.stream = decode_text(code_page, data);
                }
            }
            0x32 => {
                if let Some(m) = cur.as_mut() {
                    m.stream = utf16(data);
                }
            }
            0x31 => {
                if let Some(m) = cur.as_mut() {
                    m.offset = u32_at(data, 0).unwrap_or(0);
                }
            }
            0x22 => {
                if let Some(m) = cur.as_mut() {
                    m.kind = ModuleKind::Class;
                }
            }
            0x2B => {
                if let Some(m) = cur.take() {
                    mods.push((m, code_page));
                }
            }
            0x10 => break,
            _ => {}
        }
        pos = next;
    }
    mods
}

pub(crate) fn read_modules(bin: &[u8]) -> Result<Vec<SourceModule>, String> {
    let cf = CompoundFile::open(bin).map_err(|e| format!("the VBA project is not readable: {e}"))?;
    let dir_raw = cf.read_path_capped("VBA/dir", MAX_STREAM as u64).map_err(|e| format!("the VBA project has no directory: {e}"))?;
    let dir = decompress(&dir_raw).ok_or("the VBA project's directory is damaged")?;
    let mut out = Vec::new();
    for (m, code_page) in parse_dir(&dir) {
        let path = format!("VBA/{}", m.stream);
        let Ok(stream) = cf.read_path_capped(&path, MAX_STREAM as u64) else { continue };
        let Some(code) = stream.get(m.offset as usize..).and_then(decompress) else { continue };
        out.push(SourceModule { name: m.name, kind: m.kind, source: decode_text(code_page, &code) });
    }
    Ok(out)
}

/// Builds the `vbaProject.bin` of a project made of `modules` (used to test the reader, and so
/// to write modules back). Source is stored in `code_page` 1252.
pub fn write_bin(modules: &[SourceModule]) -> Result<Vec<u8>, String> {
    let mut dir: Vec<u8> = Vec::new();
    let push = |dir: &mut Vec<u8>, id: u16, data: &[u8]| {
        dir.extend_from_slice(&id.to_le_bytes());
        dir.extend_from_slice(&(data.len() as u32).to_le_bytes());
        dir.extend_from_slice(data);
    };
    push(&mut dir, 0x03, &1252u16.to_le_bytes());
    for m in modules {
        let name: Vec<u8> = m.name.chars().map(|c| if c.is_ascii() { c as u8 } else { b'?' }).collect();
        push(&mut dir, 0x19, &name);
        push(&mut dir, 0x1A, &name);
        let uni: Vec<u8> = m.name.encode_utf16().flat_map(u16::to_le_bytes).collect();
        push(&mut dir, 0x32, &uni);
        push(&mut dir, 0x31, &0u32.to_le_bytes());
        push(&mut dir, if m.kind == ModuleKind::Class { 0x22 } else { 0x21 }, &[]);
        push(&mut dir, 0x2B, &[]);
    }
    push(&mut dir, 0x10, &[]);
    let mut streams: Vec<(String, Vec<u8>)> = vec![("VBA/dir".to_string(), compress_literals(&dir))];
    for m in modules {
        let bytes: Vec<u8> = m.source.chars().map(|c| u8::try_from(u32::from(c)).unwrap_or(b'?')).collect();
        streams.push((format!("VBA/{}", m.name), compress_literals(&bytes)));
    }
    let mut w = nodalia_cfb::Writer::new(nodalia_cfb::Version::V3);
    for (path, data) in streams {
        w.stream_at(&path, data).map_err(|e| e.to_string())?;
    }
    Ok(w.finish())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn literal_compression_round_trips() {
        let data: Vec<u8> = (0..20_000u32).map(|i| (i % 251) as u8).collect();
        assert_eq!(decompress(&compress_literals(&data)).unwrap(), data);
        assert_eq!(decompress(&compress_literals(b"")).unwrap(), b"");
    }

    #[test]
    fn copy_tokens_expand() {
        // "abcabcabc": three literals, then a copy of 6 bytes at offset 3.
        let mut chunk = vec![0b0000_1000u8, b'a', b'b', b'c'];
        let done = 3u32;
        let bit_count = 4u32;
        let length_mask = 0xFFFFu16 >> bit_count;
        let token = (((3 - 1) as u16) << (16 - bit_count)) | ((6 - 3) as u16 & length_mask);
        let _ = done;
        chunk.extend_from_slice(&token.to_le_bytes());
        let header = 0xB000u16 | ((chunk.len() + 2 - 3) as u16);
        let mut data = vec![1u8];
        data.extend_from_slice(&header.to_le_bytes());
        data.extend_from_slice(&chunk);
        assert_eq!(decompress(&data).unwrap(), b"abcabcabc");
    }

    #[test]
    fn a_project_round_trips_through_the_file_format() {
        let modules = vec![
            SourceModule { name: "Module1".into(), kind: ModuleKind::Standard, source: "Sub A()\r\nEnd Sub\r\n".into() },
            SourceModule { name: "Sheet1".into(), kind: ModuleKind::Class, source: "Option Explicit\r\n".into() },
        ];
        let bin = write_bin(&modules).unwrap();
        let project = Project::from_bin(&bin).unwrap();
        assert_eq!(project.modules.len(), 2);
        assert_eq!(project.modules[0].source, "Sub A()\r\nEnd Sub\r\n");
        assert_eq!(project.modules[1].kind, ModuleKind::Class);
        assert_eq!(project.procedures().len(), 1);
    }

    #[test]
    fn hostile_input_is_refused_not_a_panic() {
        assert!(Project::from_bin(b"not a compound file").is_err());
        assert!(decompress(&[1, 0xFF, 0xFF]).is_none() || decompress(&[1, 0xFF, 0xFF]).is_some());
        assert!(decompress(&[2]).is_none());
    }
}
