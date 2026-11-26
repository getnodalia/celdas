//! The language services of the VBA editor, as commands: what is wrong with the code, what to offer
//! after a dot, what a name is, where it is declared, how to lay it out. They read the text the
//! editor holds (the code of a module, without its `Attribute` lines), which is rarely source that
//! compiles, so none of them needs the parser to succeed. The names of the other modules and the
//! Excel object model come from the workbook and from `excel.catalog`.

use std::sync::OnceLock;

use celdas_vba::ast::ProcKind;
use celdas_vba::ide::{self, Catalog, Severity};
use serde_json::{Value as Json, json};

use super::project_of;
use crate::cmd::str_param;
use crate::{EngineError, Result, Session};

/// The Excel object model, as the macro host implements it.
const EXCEL: &str = include_str!("excel.catalog");
/// The most text one request analyses.
const MAX_CODE: usize = 2 * 1024 * 1024;

/// The VBA library and the Excel object model.
pub fn catalog() -> &'static Catalog {
    static CATALOG: OnceLock<Catalog> = OnceLock::new();
    CATALOG.get_or_init(|| {
        let mut c = Catalog::vba();
        c.extend(EXCEL);
        c
    })
}

fn bad(msg: impl Into<String>) -> EngineError {
    crate::cmd::bad("vba", msg)
}

fn code_of(p: &Json) -> Result<&str> {
    let code = p.get("code").and_then(Json::as_str).ok_or_else(|| bad("missing `code`: the text of the module"))?;
    if code.len() > MAX_CODE {
        return Err(bad("`code` is too long to analyse"));
    }
    Ok(code)
}

/// 1-based line and 0-based column of the caret.
fn position(p: &Json) -> Result<(u32, usize)> {
    let line = p.get("line").and_then(Json::as_u64).filter(|l| *l >= 1).and_then(|l| u32::try_from(l).ok());
    let col = p.get("col").and_then(Json::as_u64).and_then(|c| usize::try_from(c).ok());
    match (line, col) {
        (Some(l), Some(c)) => Ok((l, c)),
        _ => Err(bad("`line` (1-based) and `col` (0-based, in characters) are needed")),
    }
}

/// The names the other modules of the project declare: their procedures, module-level variables,
/// types and enumeration members, and the modules' own names.
fn names_elsewhere(s: &Session, module: Option<&str>) -> Vec<String> {
    let Ok(Some((project, _))) = project_of(s) else { return Vec::new() };
    let mut out = Vec::new();
    for m in &project.modules {
        out.push(m.name.clone());
        if module.is_some_and(|n| n.eq_ignore_ascii_case(&m.name)) {
            continue;
        }
        let scan = ide::scan(&m.source);
        out.extend(scan.procs.iter().filter(|p| p.public).map(|p| p.name.clone()));
        out.extend(scan.vars.iter().filter(|v| v.proc.is_none() && v.public).map(|v| v.name.clone()));
        for t in &scan.types {
            out.push(t.name.clone());
            if t.is_enum {
                out.extend(t.fields.iter().map(|(f, _)| f.clone()));
            }
        }
    }
    out
}

fn severity(s: Severity) -> &'static str {
    match s {
        Severity::Error => "error",
        Severity::Warning => "warning",
        Severity::Hint => "hint",
    }
}

fn kind_name(k: ProcKind) -> &'static str {
    match k {
        ProcKind::Sub => "sub",
        ProcKind::Function => "function",
        ProcKind::PropertyGet => "propertyGet",
        ProcKind::PropertyLet => "propertyLet",
        ProcKind::PropertySet => "propertySet",
    }
}

fn item_kind(k: ide::complete::ItemKind) -> &'static str {
    use ide::complete::ItemKind as K;
    match k {
        K::Keyword => "keyword",
        K::Snippet => "snippet",
        K::Function => "function",
        K::Constant => "constant",
        K::Variable => "variable",
        K::Property => "property",
        K::Method => "method",
        K::Class => "class",
        K::Procedure => "procedure",
        K::Type => "type",
    }
}

/// `vba.analyze {code, module?}` → the marks, the procedures, the variables and the types.
pub fn analyze(s: &mut Session, p: &Json) -> Result<Json> {
    let code = code_of(p)?;
    let module = str_param(p, "module").unwrap_or("Module1");
    let extra = names_elsewhere(s, Some(module));
    let a = ide::analyze(module, code, catalog(), &extra);
    let diagnostics: Vec<Json> = a
        .diagnostics
        .iter()
        .map(|d| {
            let args: serde_json::Map<String, Json> = d.args.iter().map(|(k, v)| ((*k).to_string(), json!(v))).collect();
            json!({"line": d.line, "start": d.start, "end": d.end, "severity": severity(d.severity), "message": d.message, "template": d.template, "args": args})
        })
        .collect();
    let procedures: Vec<Json> = a
        .scan
        .procs
        .iter()
        .map(|p| {
            json!({"name": p.name, "kind": kind_name(p.kind), "public": p.public, "line": p.line, "endLine": p.end_line, "col": p.name_col, "signature": p.signature(), "doc": p.doc})
        })
        .collect();
    let variables: Vec<Json> = a
        .scan
        .vars
        .iter()
        .map(|v| {
            let procedure = v.proc.and_then(|i| a.scan.procs.get(i)).map(|p| p.name.clone());
            json!({"name": v.name, "type": v.ty, "line": v.line, "col": v.col, "procedure": procedure, "param": v.param, "constant": v.constant, "public": v.public})
        })
        .collect();
    let types: Vec<Json> = a.scan.types.iter().map(|t| json!({"name": t.name, "enum": t.is_enum, "line": t.line})).collect();
    Ok(
        json!({"diagnostics": diagnostics, "procedures": procedures, "variables": variables, "types": types, "optionExplicit": a.scan.option_explicit}),
    )
}

/// `vba.complete {code, line, col, explicit?, module?}` → what to offer, and the call being written.
pub fn complete(s: &mut Session, p: &Json) -> Result<Json> {
    let code = code_of(p)?;
    let (line, col) = position(p)?;
    let explicit = p.get("explicit").and_then(Json::as_bool).unwrap_or(false);
    let _ = s;
    let scan = ide::scan(code);
    let c = ide::complete::complete(&scan, catalog(), code, line, col, explicit);
    let items: Vec<Json> =
        c.items.iter().map(|i| json!({"label": i.label, "kind": item_kind(i.kind), "detail": i.detail, "doc": i.doc, "insert": i.insert})).collect();
    let signature = ide::complete::signature_help(&scan, catalog(), code, line, col)
        .map(|g| json!({"label": g.label, "params": g.params, "active": g.active, "doc": g.doc}));
    Ok(json!({"from": c.from, "items": items, "signature": signature}))
}

/// `vba.hover {code, line, col}` → `{signature, doc, definition?: {line, col}}` for the name there.
pub fn hover(_: &mut Session, p: &Json) -> Result<Json> {
    let code = code_of(p)?;
    let (line, col) = position(p)?;
    let scan = ide::scan(code);
    let Some((signature, doc)) = ide::complete::hover(&scan, catalog(), code, line, col) else { return Ok(json!({"signature": null})) };
    let definition = ide::complete::definition(&scan, line, col).map(|(l, c)| json!({"line": l, "col": c}));
    Ok(json!({"signature": signature, "doc": doc, "definition": definition}))
}

/// `vba.occurrences {code, line, col}` → every place the name under the caret is written.
pub fn occurrences(_: &mut Session, p: &Json) -> Result<Json> {
    let code = code_of(p)?;
    let (line, col) = position(p)?;
    let scan = ide::scan(code);
    let Some((name, ..)) = ide::complete::word_at(&scan, line, col) else { return Ok(json!({"name": null, "places": []})) };
    let places: Vec<Json> = ide::complete::occurrences(&scan, &name).into_iter().map(|(l, a, b)| json!({"line": l, "start": a, "end": b})).collect();
    let definition = ide::complete::definition(&scan, line, col).map(|(l, c)| json!({"line": l, "col": c}));
    Ok(json!({"name": name, "places": places, "definition": definition}))
}

/// `vba.format {code}` → the code laid out: indented by block, the language's words spelled the
/// usual way, no trailing blanks.
pub fn format(_: &mut Session, p: &Json) -> Result<Json> {
    let code = code_of(p)?;
    let out = ide::format(code);
    // The file keeps `\r\n`; an editor that sent them gets them back.
    let out = if code.contains("\r\n") { out.replace("\r\n", "\n").replace('\n', "\r\n") } else { out };
    Ok(json!({"changed": out != code, "code": out}))
}

#[cfg(test)]
mod tests {
    use serde_json::json;

    use crate::Session;

    fn session() -> Session {
        let mut s = Session::new();
        s.new_workbook();
        s
    }

    #[test]
    fn the_catalog_knows_the_excel_constants_the_host_runs() {
        let c = super::catalog();
        for name in ["xlCenter", "xlUp", "xlCalculationManual", "vbCrLf"] {
            assert!(c.global(name).is_some(), "{name}");
        }
        let text = include_str!("excel.catalog");
        for line in text.lines().filter_map(|l| l.strip_prefix("const ")) {
            let name = line.split_whitespace().next().unwrap_or("").to_lowercase();
            assert!(super::super::consts::excel_constant(&name).is_some(), "{name} is not a constant the host knows");
        }
        // Every constant the host knows is listed, so the editor offers all of them.
        let consts = include_str!("consts.rs");
        let body = consts.split("#[cfg(test)]").next().unwrap_or("");
        for token in body.split('"').skip(1).step_by(2).filter(|t| t.starts_with("xl")) {
            assert!(
                text.lines()
                    .any(|l| l.strip_prefix("const ").and_then(|r| r.split_whitespace().next()).is_some_and(|n| n.eq_ignore_ascii_case(token))),
                "{token} is missing from excel.catalog"
            );
        }
    }

    #[test]
    fn analyze_marks_a_missing_end_and_lists_the_procedures() {
        let mut s = session();
        let r = s.execute("vba.analyze", json!({"code": "Sub A()\n    Dim x As Long\n    If x Then\nEnd Sub\n"})).unwrap();
        let first = &r["diagnostics"][0];
        assert_eq!(first["severity"], "error", "{r}");
        assert_eq!(r["procedures"][0]["name"], "A");
        assert_eq!(r["procedures"][0]["signature"], "Sub A()");
        assert!(r["variables"].as_array().unwrap().iter().any(|v| v["name"] == "x" && v["procedure"] == "A"));
    }

    #[test]
    fn complete_follows_the_excel_object_model() {
        let mut s = session();
        let code = "Sub A()\n    Dim ws As Worksheet\n    ws.Range(\"A1\").Fo\nEnd Sub\n";
        let r = s.execute("vba.complete", json!({"code": code, "line": 3, "col": 22})).unwrap();
        let labels: Vec<&str> = r["items"].as_array().unwrap().iter().map(|i| i["label"].as_str().unwrap()).collect();
        assert!(labels.contains(&"Font") && labels.contains(&"Formula"), "{labels:?}");
        let r = s.execute("vba.complete", json!({"code": "Sub A()\n    x = Mid(\"abc\", \nEnd Sub\n", "line": 2, "col": 19})).unwrap();
        assert_eq!(r["signature"]["active"], 1);
    }

    #[test]
    fn hover_definition_occurrences_and_format() {
        let mut s = session();
        let code = "Sub A()\n    Dim total As Long\n    total = 1\nEnd Sub\n";
        let h = s.execute("vba.hover", json!({"code": code, "line": 3, "col": 6})).unwrap();
        assert_eq!(h["signature"], "total As Long");
        assert_eq!(h["definition"]["line"], 2);
        let o = s.execute("vba.occurrences", json!({"code": code, "line": 3, "col": 6})).unwrap();
        assert_eq!(o["places"].as_array().unwrap().len(), 2);
        let f = s.execute("vba.format", json!({"code": "sub a()\r\ndim x as long\r\nend sub\r\n"})).unwrap();
        assert_eq!(f["code"], "Sub a()\r\n    Dim x As Long\r\nEnd Sub\r\n");
        assert!(s.execute("vba.complete", json!({"code": "x"})).is_err());
        assert!(s.execute("vba.analyze", json!({})).is_err());
    }

    #[test]
    fn a_workbook_without_a_project_can_be_given_one() {
        let mut s = session();
        let r = s.execute("vba.create", json!({"code": "Sub Hola()\n    Cells(1, 1).Value = 7\nEnd Sub\n"})).unwrap();
        assert_eq!(r["module"], "Module1");
        assert!(s.execute("vba.create", json!({})).is_err());
        s.execute("vba.run", json!({"procedure": "Hola"})).unwrap();
        let src = s.execute("vba.source", json!({})).unwrap();
        let module1 = src["modules"].as_array().unwrap().iter().find(|m| m["name"] == "Module1").unwrap();
        assert!(module1["code"].as_str().unwrap().contains("Sub Hola()"));
        // The project starts with the code behind the workbook and the sheet, too.
        assert_eq!(src["modules"][0]["name"], "ThisWorkbook");
        assert_eq!(src["modules"][0]["document"], true);
    }
}
