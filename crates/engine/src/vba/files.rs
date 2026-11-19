//! The files a macro opens with `Open … For Output`, `Print #`, `Line Input #`, `Dir`, `MkDir`
//! and `Kill`. A file written is held in memory and goes to disk, whole and atomically, when the
//! macro closes it, so a macro that fails half way leaves no half-written file behind.

use std::collections::{HashMap, VecDeque};

use celdas_vba::ast::FileMode;

/// The path as the platform spells it: a macro written on Windows uses backslashes.
pub(super) fn norm(path: &str) -> String {
    if cfg!(windows) { path.to_string() } else { path.replace('\\', "/") }
}

enum OpenFile {
    Out { path: String, text: String },
    In { lines: VecDeque<String> },
}

#[derive(Default)]
pub struct Files {
    open: HashMap<i32, OpenFile>,
    listing: VecDeque<String>,
}

/// How many bytes one file a macro writes may hold.
const MAX_FILE_BYTES: usize = 256 * 1024 * 1024;

fn text_of(bytes: &[u8]) -> String {
    match std::str::from_utf8(bytes) {
        Ok(t) => t.to_string(),
        // Files an older macro wrote in the Windows code page: every byte is a character.
        Err(_) => bytes.iter().map(|&b| char::from(b)).collect(),
    }
}

impl Files {
    pub fn open(&mut self, path: &str, mode: FileMode, number: i32) -> Result<(), String> {
        if self.open.contains_key(&number) {
            return Err("File already open".into());
        }
        let path = norm(path);
        let file = match mode {
            FileMode::Output => OpenFile::Out { path, text: String::new() },
            FileMode::Append => {
                let text = match crate::io::read_file(&path) {
                    Ok(b) => text_of(&b),
                    Err(_) => String::new(),
                };
                OpenFile::Out { path, text }
            }
            FileMode::Input => {
                let bytes = crate::io::read_file(&path).map_err(|e| format!("Bad file name or number: {e}"))?;
                let text = text_of(&bytes);
                OpenFile::In { lines: text.lines().map(str::to_string).collect() }
            }
            FileMode::Binary | FileMode::Random => return Err("Only text files can be opened".into()),
        };
        // A write that can't land (a folder that doesn't exist, no permission) fails at Open, as it does in Excel.
        if let OpenFile::Out { path, text } = &file
            && text.is_empty()
        {
            crate::io::write_file(path, b"").map_err(|e| e.to_string())?;
        }
        self.open.insert(number, file);
        Ok(())
    }

    pub fn write(&mut self, number: i32, text: &str) -> Result<(), String> {
        match self.open.get_mut(&number) {
            Some(OpenFile::Out { text: buf, .. }) => {
                if buf.len().saturating_add(text.len()) > MAX_FILE_BYTES {
                    return Err("Disk full".into());
                }
                buf.push_str(text);
                Ok(())
            }
            Some(OpenFile::In { .. }) => Err("Bad file mode".into()),
            None => Err("Bad file name or number".into()),
        }
    }

    pub fn read_line(&mut self, number: i32) -> Result<Option<String>, String> {
        match self.open.get_mut(&number) {
            Some(OpenFile::In { lines }) => Ok(lines.pop_front()),
            Some(OpenFile::Out { .. }) => Err("Bad file mode".into()),
            None => Err("Bad file name or number".into()),
        }
    }

    pub fn at_eof(&self, number: i32) -> Result<bool, String> {
        match self.open.get(&number) {
            Some(OpenFile::In { lines }) => Ok(lines.is_empty()),
            Some(OpenFile::Out { .. }) => Err("Bad file mode".into()),
            None => Err("Bad file name or number".into()),
        }
    }

    pub fn close(&mut self, number: i32, written: &mut Vec<String>) -> Result<(), String> {
        match self.open.remove(&number) {
            Some(OpenFile::Out { path, text }) => {
                crate::io::write_file(&path, text.as_bytes()).map_err(|e| e.to_string())?;
                written.push(path);
                Ok(())
            }
            _ => Ok(()),
        }
    }

    pub fn close_all(&mut self, written: &mut Vec<String>) -> Result<(), String> {
        let mut numbers: Vec<i32> = self.open.keys().copied().collect();
        numbers.sort_unstable();
        let mut first_error = None;
        for n in numbers {
            if let Err(e) = self.close(n, written) {
                first_error.get_or_insert(e);
            }
        }
        first_error.map_or(Ok(()), Err)
    }

    /// `Dir(pattern, attributes)`: the first match, or with no pattern the next one.
    pub fn dir(&mut self, pattern: &str, attributes: i32) -> Result<String, String> {
        if pattern.is_empty() {
            return Ok(self.listing.pop_front().unwrap_or_default());
        }
        self.listing = list(&norm(pattern), attributes & 16 != 0).into();
        Ok(self.listing.pop_front().unwrap_or_default())
    }
}

/// Whether `name` matches `pattern`, where `*` is any run of characters and `?` any one; case does not matter.
fn glob(pattern: &[char], name: &[char]) -> bool {
    match pattern.split_first() {
        None => name.is_empty(),
        Some(('*', rest)) => (0..=name.len()).any(|i| glob(rest, name.get(i..).unwrap_or(&[]))),
        Some((p, rest)) => name.split_first().is_some_and(|(n, tail)| (*p == '?' || p.to_lowercase().eq(n.to_lowercase())) && glob(rest, tail)),
    }
}

#[cfg(not(target_arch = "wasm32"))]
fn list(pattern: &str, directories: bool) -> Vec<String> {
    let (dir, file) = pattern.rsplit_once('/').map_or((".", pattern), |(d, f)| (if d.is_empty() { "/" } else { d }, f));
    if file.is_empty() {
        return if directories && std::path::Path::new(dir).is_dir() { vec![".".into()] } else { Vec::new() };
    }
    let pat: Vec<char> = file.chars().collect();
    let mut names: Vec<String> = std::fs::read_dir(dir)
        .map(|rd| {
            rd.filter_map(|e| e.ok())
                .filter(|e| e.file_type().is_ok_and(|t| t.is_file() || (directories && t.is_dir())))
                .filter_map(|e| e.file_name().into_string().ok())
                .filter(|n| glob(&pat, &n.chars().collect::<Vec<_>>()))
                .collect()
        })
        .unwrap_or_default();
    names.sort_unstable();
    names
}

#[cfg(target_arch = "wasm32")]
fn list(_pattern: &str, _directories: bool) -> Vec<String> {
    Vec::new()
}

#[cfg(not(target_arch = "wasm32"))]
pub fn make_dir(path: &str) -> Result<(), String> {
    let path = norm(path);
    // `MkDir` needs the parent to exist and refuses a folder that is there already.
    std::fs::create_dir(&path).map_err(|e| match e.kind() {
        std::io::ErrorKind::AlreadyExists => "Path/File access error: the folder already exists".to_string(),
        _ => format!("Path not found: {path}"),
    })
}

#[cfg(target_arch = "wasm32")]
pub fn make_dir(path: &str) -> Result<(), String> {
    Err(format!("{path}: no file system in the browser"))
}

#[cfg(not(target_arch = "wasm32"))]
pub fn remove_file(path: &str) -> Result<(), String> {
    let path = norm(path);
    std::fs::remove_file(&path).map_err(|_| format!("File not found: {path}"))
}

#[cfg(target_arch = "wasm32")]
pub fn remove_file(path: &str) -> Result<(), String> {
    Err(format!("{path}: no file system in the browser"))
}

#[cfg(not(target_arch = "wasm32"))]
pub fn environ(name: &str) -> String {
    std::env::var(name).unwrap_or_default()
}

#[cfg(target_arch = "wasm32")]
pub fn environ(_name: &str) -> String {
    String::new()
}

/// The user's Desktop folder (`WScript.Shell.SpecialFolders("Desktop")`), or the home folder when there is none.
pub fn desktop() -> String {
    let home = if cfg!(windows) { environ("USERPROFILE") } else { environ("HOME") };
    if home.is_empty() {
        return String::new();
    }
    let desktop = format!("{home}{}Desktop", if cfg!(windows) { '\\' } else { '/' });
    #[cfg(not(target_arch = "wasm32"))]
    if !std::path::Path::new(&desktop).is_dir() {
        return home;
    }
    desktop
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn glob_matches_like_dir() {
        let g = |p: &str, n: &str| glob(&p.chars().collect::<Vec<_>>(), &n.chars().collect::<Vec<_>>());
        assert!(g("*.txt", "Datos.TXT"));
        assert!(g("Rebanada_??.xlsx", "Rebanada_01.xlsx"));
        assert!(!g("*.txt", "a.xlsx"));
        assert!(g("*", "anything"));
    }

    #[test]
    fn a_written_file_lands_when_it_is_closed() {
        let dir = std::env::temp_dir().join(format!("celdas-vba-files-{}", std::process::id()));
        let _ = std::fs::create_dir_all(&dir);
        let path = dir.join("out.txt").to_string_lossy().to_string();
        let mut f = Files::default();
        let mut written = Vec::new();
        f.open(&path, FileMode::Output, 1).unwrap();
        f.write(1, "a\r\n").unwrap();
        f.write(1, "b\r\n").unwrap();
        assert_eq!(std::fs::read_to_string(&path).unwrap(), "");
        f.close(1, &mut written).unwrap();
        assert_eq!(std::fs::read_to_string(&path).unwrap(), "a\r\nb\r\n");
        assert_eq!(written, vec![path.clone()]);
        f.open(&path, FileMode::Input, 2).unwrap();
        assert_eq!(f.read_line(2).unwrap().as_deref(), Some("a"));
        assert_eq!(f.read_line(2).unwrap().as_deref(), Some("b"));
        assert_eq!(f.read_line(2).unwrap(), None);
        assert_eq!(f.dir(&format!("{}/*.txt", dir.to_string_lossy()), 0).unwrap(), "out.txt");
        assert_eq!(f.dir("", 0).unwrap(), "");
        let _ = std::fs::remove_dir_all(&dir);
    }
}
