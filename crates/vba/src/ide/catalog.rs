// SPDX-License-Identifier: MIT OR Apache-2.0
//! The names an editor offers besides the keywords: functions, constants, global objects and the
//! members of each class, with their signatures. A catalog is written as text, one line per name:
//!
//! ```text
//! class Range item=Range "A block of cells"
//!   p Value : Variant "What the cells hold"
//!   p Address([RowAbsolute]) : String ro "The A1 reference"
//!   m Offset(RowOffset, ColumnOffset) : Range "Moves the range"
//! fn MsgBox(Prompt, [Buttons]) : Long "Shows a message"
//! const vbCrLf : String "Carriage return and line feed"
//! var ActiveSheet : Worksheet "The sheet in front"
//! ```
//!
//! `p` is a property (`ro` when it cannot be set), `m` a method, `fn` a global function, `const` a
//! constant and `var` a global property or object. A name may take `(parameters)`, a type after
//! ` : ` and a quoted description. The VBA language's own names are [`Catalog::vba`]; the host
//! adds its object model with [`Catalog::extend`].

use std::collections::HashMap;

/// Whether a member is read and written like a variable or called.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum MemberKind {
    Property,
    Method,
}

/// A property or method of a class.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Member {
    pub name: String,
    pub kind: MemberKind,
    /// The parameters as written (`Cell1, [Cell2]`), without the parentheses.
    pub params: String,
    /// The class or type the member answers with; empty for none or for a plain value type.
    pub returns: String,
    pub readonly: bool,
    pub doc: String,
}

impl Member {
    /// `Name(Cell1, [Cell2]) As Range`, the text of a signature tip.
    pub fn signature(&self) -> String {
        let mut s = self.name.clone();
        if !self.params.is_empty() {
            s.push('(');
            s.push_str(&self.params);
            s.push(')');
        }
        if !self.returns.is_empty() {
            s.push_str(" As ");
            s.push_str(&self.returns);
        }
        s
    }
}

/// A class of objects with its members.
#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct Class {
    pub name: String,
    pub doc: String,
    pub members: Vec<Member>,
    /// What `obj(index)` and `For Each` give: the class of its items.
    pub item: Option<String>,
}

impl Class {
    /// A member by name, ignoring case.
    pub fn member(&self, name: &str) -> Option<&Member> {
        self.members.iter().find(|m| m.name.eq_ignore_ascii_case(name))
    }
}

/// What a global name is.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum GlobalKind {
    Function,
    Constant,
    /// A global property or object (`ActiveSheet`, `Application`).
    Variable,
}

/// A name that needs no declaration.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Global {
    pub name: String,
    pub kind: GlobalKind,
    pub params: String,
    pub returns: String,
    pub doc: String,
}

impl Global {
    pub fn signature(&self) -> String {
        let mut s = self.name.clone();
        if self.kind == GlobalKind::Function && !self.params.is_empty() {
            s.push('(');
            s.push_str(&self.params);
            s.push(')');
        }
        if !self.returns.is_empty() {
            s.push_str(" As ");
            s.push_str(&self.returns);
        }
        s
    }
}

/// Classes and global names, found by name ignoring case.
#[derive(Clone, Debug, Default)]
pub struct Catalog {
    pub classes: Vec<Class>,
    pub globals: Vec<Global>,
    class_index: HashMap<String, usize>,
    global_index: HashMap<String, usize>,
}

const VBA_CATALOG: &str = include_str!("vba.catalog");

impl Catalog {
    /// The names of the VBA language itself.
    pub fn vba() -> Catalog {
        Catalog::parse(VBA_CATALOG)
    }

    /// Reads the catalog text. A line it cannot read is skipped.
    pub fn parse(text: &str) -> Catalog {
        let mut c = Catalog::default();
        c.extend(text);
        c
    }

    /// Adds the names of more catalog text.
    pub fn extend(&mut self, text: &str) {
        let mut open: Option<usize> = None;
        for raw in text.lines() {
            let line = raw.trim();
            if line.is_empty() || line.starts_with('#') {
                continue;
            }
            let indented = raw.starts_with(char::is_whitespace);
            let (word, rest) = line.split_once(char::is_whitespace).unwrap_or((line, ""));
            match word {
                "class" => open = self.add_class(rest),
                "p" | "m" if indented => {
                    if let Some(i) = open
                        && let Some(class) = self.classes.get_mut(i)
                    {
                        class.members.push(parse_member(word == "m", rest));
                    }
                }
                "fn" | "const" | "var" => {
                    open = None;
                    self.add_global(word, rest);
                }
                _ => {}
            }
        }
    }

    fn add_class(&mut self, rest: &str) -> Option<usize> {
        let (head, doc) = split_doc(rest);
        let mut words = head.split_whitespace();
        let name = words.next()?.to_string();
        let item = words.find_map(|w| w.strip_prefix("item=")).map(str::to_string);
        let key = name.to_lowercase();
        if let Some(&i) = self.class_index.get(&key) {
            // A later text may add members to a class.
            if let Some(class) = self.classes.get_mut(i) {
                if class.item.is_none() {
                    class.item = item;
                }
                if class.doc.is_empty() {
                    class.doc = doc;
                }
            }
            return Some(i);
        }
        self.classes.push(Class { name, doc, members: Vec::new(), item });
        self.class_index.insert(key, self.classes.len() - 1);
        Some(self.classes.len() - 1)
    }

    fn add_global(&mut self, word: &str, rest: &str) {
        let (head, doc) = split_doc(rest);
        let (decl, _) = split_returns(&head);
        let (name, params) = split_name(decl);
        let (_, returns) = split_returns(&head);
        if name.is_empty() {
            return;
        }
        let kind = match word {
            "fn" => GlobalKind::Function,
            "const" => GlobalKind::Constant,
            _ => GlobalKind::Variable,
        };
        let g = Global { name: name.to_string(), kind, params: params.to_string(), returns: returns.trim_end_matches(" ro").to_string(), doc };
        let key = g.name.to_lowercase();
        match self.global_index.get(&key) {
            Some(&i) => {
                if let Some(slot) = self.globals.get_mut(i) {
                    *slot = g;
                }
            }
            None => {
                self.global_index.insert(key, self.globals.len());
                self.globals.push(g);
            }
        }
    }

    /// A class by name, ignoring case.
    pub fn class(&self, name: &str) -> Option<&Class> {
        self.class_index.get(&name.to_lowercase()).and_then(|&i| self.classes.get(i))
    }

    /// A global name, ignoring case.
    pub fn global(&self, name: &str) -> Option<&Global> {
        self.global_index.get(&name.to_lowercase()).and_then(|&i| self.globals.get(i))
    }
}

/// The text before a trailing `"description"`, and the description.
fn split_doc(rest: &str) -> (String, String) {
    let rest = rest.trim();
    if let Some(start) = rest.find('"')
        && let Some(end) = rest.rfind('"')
        && end > start
    {
        let doc = rest.get(start + 1..end).unwrap_or("").to_string();
        return (rest.get(..start).unwrap_or("").trim().to_string(), doc);
    }
    (rest.to_string(), String::new())
}

/// `Name(args)` and the text after ` : `.
fn split_returns(head: &str) -> (&str, &str) {
    match head.rsplit_once(" : ") {
        Some((decl, ty)) => (decl.trim(), ty.trim()),
        None => (head.trim(), ""),
    }
}

/// `Name` and the parameters of `Name(parameters)`.
fn split_name(decl: &str) -> (&str, &str) {
    match decl.split_once('(') {
        Some((name, rest)) => (name.trim(), rest.rsplit_once(')').map_or(rest, |(p, _)| p).trim()),
        None => (decl.trim(), ""),
    }
}

fn parse_member(method: bool, rest: &str) -> Member {
    let (head, doc) = split_doc(rest);
    let (decl, ty) = split_returns(&head);
    let readonly = ty.ends_with(" ro") || (ty.is_empty() && decl.ends_with(" ro"));
    let ty = ty.trim_end_matches(" ro").trim();
    let decl = decl.trim_end_matches(" ro").trim();
    let (name, params) = split_name(decl);
    Member {
        name: name.to_string(),
        kind: if method { MemberKind::Method } else { MemberKind::Property },
        params: params.to_string(),
        returns: ty.to_string(),
        readonly,
        doc,
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn a_catalog_line_is_read_into_its_parts() {
        let c = Catalog::parse(
            "class Range item=Range \"A block of cells\"\n  p Value : Variant \"What it holds\"\n  p Address([A]) : String ro \"The reference\"\n  m Offset(R, C) : Range \"Moves\"\n  m Select\nfn MsgBox(Prompt, [Buttons]) : Long \"Shows\"\nconst vbCrLf : String \"New line\"\nvar ActiveSheet : Worksheet \"The sheet\"\n",
        );
        let range = c.class("RANGE").unwrap();
        assert_eq!(range.item.as_deref(), Some("Range"));
        assert_eq!(range.members.len(), 4);
        let address = range.member("address").unwrap();
        assert!(address.readonly);
        assert_eq!(address.signature(), "Address([A]) As String");
        assert_eq!(range.member("Select").unwrap().kind, MemberKind::Method);
        assert_eq!(c.global("msgbox").unwrap().signature(), "MsgBox(Prompt, [Buttons]) As Long");
        assert_eq!(c.global("VBCRLF").unwrap().kind, GlobalKind::Constant);
        assert_eq!(c.global("activesheet").unwrap().returns, "Worksheet");
    }

    #[test]
    fn the_language_catalog_reads_and_has_the_common_names() {
        let c = Catalog::vba();
        for name in ["Left", "MsgBox", "UBound", "vbYes", "Format", "InStr"] {
            assert!(c.global(name).is_some(), "{name}");
        }
        assert!(c.class("Dictionary").unwrap().member("Exists").is_some());
        assert!(c.class("Err").unwrap().member("Number").is_some());
        assert!(c.globals.iter().all(|g| !g.name.is_empty()));
    }

    /// The names the interpreter knows and the names an editor offers must not drift apart.
    #[test]
    fn every_constant_the_interpreter_knows_is_in_the_catalog() {
        let src = include_str!("../builtins.rs");
        let body = src.split("pub fn vb_constant").nth(1).and_then(|s| s.split("fn opt(").next()).unwrap_or("");
        let c = Catalog::vba();
        let mut checked = 0;
        for part in body.split('"').skip(1).step_by(2) {
            if part.starts_with("vb") && part.chars().all(|ch| ch.is_ascii_alphanumeric()) {
                assert!(c.global(part).is_some(), "{part} is a constant of the interpreter but not of the catalog");
                checked += 1;
            }
        }
        assert!(checked > 40, "{checked}");
        for f in ["ChDrive", "LenB", "Environ", "Randomize", "Timer", "Weekday"] {
            assert!(c.global(f).is_some(), "{f}");
        }
    }

    #[test]
    fn nonsense_lines_are_skipped() {
        let c = Catalog::parse("???\n  p orphan\nclass\nfn\nm stray\n\"\"");
        assert!(c.classes.is_empty());
        assert!(c.globals.is_empty());
    }

    #[test]
    fn a_class_can_be_extended_by_a_later_text() {
        let mut c = Catalog::parse("class A\n  p X : Long\n");
        c.extend("class A\n  p Y : Long\n");
        assert_eq!(c.class("a").unwrap().members.len(), 2);
    }
}
