//! Running the VBA project a workbook carries (`vbaProject.bin`).
//!
//! The language is [`celdas_vba`]'s; this module is the Excel object model behind it ([`ExcelHost`])
//! and the three commands that list, show and run the project's procedures. A run is one undo step
//! and all or nothing: when the macro fails, the workbook, and every file it would have written,
//! stay as they were.

mod app;
mod consts;
mod files;
mod host;
mod objects;
mod range;
mod shapes;

use std::hash::{Hash, Hasher};

use celdas_vba::{Project, SourceModule, Vm};
use serde_json::{Value as Json, json};

use crate::cmd::{CommandSpec, bool_param, cmd, has_doc, str_param};
use crate::{EngineError, Result, Session, Trust};

pub use host::{ExcelHost, Message};

pub fn specs() -> Vec<CommandSpec> {
    vec![
        cmd!(
            query "vba.list",
            "List VBA Procedures",
            [],
            None,
            "{} → {present, modules: [{name, kind: standard|class, lines}], procedures: [{module, name, kind: sub|function, public, takesArguments}], errors: [{module, message}]}",
            has_doc,
            list
        ),
        cmd!(
            query "vba.source",
            "Show VBA Source",
            [],
            None,
            "{module?: name} → {modules: [{name, kind, source (the whole text), header (its leading Attribute lines), code (the rest: what an editor shows), document (the code behind ThisWorkbook or a sheet)}]}",
            has_doc,
            source
        ),
        cmd!(
            "vba.save",
            "Save VBA Module",
            [],
            None,
            "{module: \"Module1\", code: \"Sub Hello()…\", kind?: standard|class (for a module the project doesn't have yet, default standard)} replaces the code of a module of the workbook's VBA project, or adds the module, and keeps the rest of the project as it is (its `Attribute` lines stay) → {module, created, procedures: [name], errors: [message]}. The code is saved even when Celdas can't read it (errors say why). One undo step; the workbook then asks again before running a macro",
            has_doc,
            save
        ),
        cmd!(
            "vba.remove",
            "Remove VBA Module",
            [],
            None,
            "{module: \"Module1\"} removes a module from the VBA project. The code behind ThisWorkbook and the sheets can't be removed (its code can be emptied with vba.save)",
            has_doc,
            remove
        ),
        cmd!(
            "vba.run",
            "Run VBA Macro",
            [],
            None,
            "{procedure: \"Name\" | \"Module.Name\", caller?: \"Shape 1\" (what Application.Caller answers: the shape that was clicked), approve?: bool} runs as one undo step; an error rolls it all back → {procedure, messages: [{title, text}], files: [paths], steps}. A workbook opened from a file asks first (an approval) because a macro can write files",
            has_doc,
            run
        ),
    ]
}

fn project_of(s: &Session) -> Result<Option<(Project, u64)>> {
    let d = s.doc()?;
    let Some(bytes) = &d.wb.vba_project else { return Ok(None) };
    let project = Project::from_bin(bytes).map_err(|e| EngineError::Other(format!("The VBA project can't be read: {e}")))?;
    let mut h = std::collections::hash_map::DefaultHasher::new();
    bytes.hash(&mut h);
    Ok(Some((project, h.finish())))
}

/// The procedures a workbook's VBA project offers to run (the rows `view.macros` adds): public
/// Subs that take no arguments, as the Macros dialog of Excel lists them. A project that can't be
/// read offers none; `vba.list` says why.
pub(crate) fn runnable_rows(wb: &celdas_model::Workbook) -> Vec<Json> {
    let Some(bytes) = &wb.vba_project else { return Vec::new() };
    let Ok(project) = Project::from_bin(bytes) else { return Vec::new() };
    project
        .procedures()
        .into_iter()
        .filter(|p| p.public && !p.is_function && !p.has_params)
        .map(|p| json!({"name": p.name, "module": p.module, "kind": "vbaProc"}))
        .collect()
}

fn list(s: &mut Session, _: &Json) -> Result<Json> {
    let Some((project, _)) = project_of(s)? else { return Ok(json!({"present": false, "modules": [], "procedures": [], "errors": []})) };
    let modules: Vec<Json> = project
        .modules
        .iter()
        .map(|m| json!({"name": m.name, "kind": kind_name(m.kind), "document": is_document(&m.source), "lines": m.source.lines().count()}))
        .collect();
    let procedures: Vec<Json> = project
        .procedures()
        .into_iter()
        .map(|p| json!({"module": p.module, "name": p.name, "kind": if p.is_function { "function" } else { "sub" }, "public": p.public, "takesArguments": p.has_params}))
        .collect();
    let errors: Vec<Json> = project.syntax_errors().into_iter().map(|(m, e)| json!({"module": m, "message": e.to_string()})).collect();
    Ok(json!({"present": true, "modules": modules, "procedures": procedures, "errors": errors}))
}

fn source(s: &mut Session, p: &Json) -> Result<Json> {
    let Some((project, _)) = project_of(s)? else { return Err(EngineError::Other("The workbook has no VBA project.".into())) };
    let only = str_param(p, "module");
    let modules: Vec<Json> = project
        .modules
        .iter()
        .filter(|m| only.is_none_or(|n| m.name.eq_ignore_ascii_case(n)))
        .map(|m| {
            let (header, code) = split_attributes(&m.source);
            json!({"name": m.name, "kind": kind_name(m.kind), "source": m.source, "header": header, "code": code, "document": is_document(&m.source)})
        })
        .collect();
    if modules.is_empty() && only.is_some() {
        return Err(EngineError::Other(format!("The VBA project has no module '{}'.", only.unwrap_or_default())));
    }
    Ok(json!({"modules": modules}))
}

fn kind_name(k: celdas_vba::ModuleKind) -> &'static str {
    if k == celdas_vba::ModuleKind::Standard { "standard" } else { "class" }
}

/// A module's leading `Attribute` lines (how the file names and declares it) and the rest, the code
/// a person edits.
fn split_attributes(source: &str) -> (&str, &str) {
    let mut end = 0;
    for line in source.split_inclusive('\n') {
        if line.trim_start().to_ascii_lowercase().starts_with("attribute ") {
            end += line.len();
        } else {
            break;
        }
    }
    source.split_at(end)
}

/// Whether the module is the code behind ThisWorkbook or a sheet (or a form): the file declares it
/// with a base class, and the project can't lose it.
fn is_document(source: &str) -> bool {
    split_attributes(source).0.to_ascii_lowercase().contains("vb_base")
}

fn project_error(e: String) -> EngineError {
    EngineError::Other(format!("The VBA project can't be written: {e}"))
}

fn modules_of(project: &Project) -> Vec<SourceModule> {
    project.modules.iter().map(|m| SourceModule { name: m.name.clone(), kind: m.kind, source: m.source.clone() }).collect()
}

fn replace_project(s: &mut Session, bytes: Vec<u8>) -> Result<()> {
    crate::cmd::edit(s, |cx| {
        cx.wb.vba_project = Some(std::sync::Arc::new(bytes));
        Ok(())
    })
}

fn save(s: &mut Session, p: &Json) -> Result<Json> {
    let name = str_param(p, "module")
        .map(str::trim)
        .filter(|n| !n.is_empty())
        .ok_or_else(|| EngineError::Other("vba.save: missing `module`".into()))?
        .to_string();
    let code = p.get("code").and_then(Json::as_str).ok_or_else(|| EngineError::Other("vba.save: missing `code`".into()))?;
    let kind = match str_param(p, "kind") {
        None | Some("standard") => celdas_vba::ModuleKind::Standard,
        Some("class") => celdas_vba::ModuleKind::Class,
        Some(other) => return Err(EngineError::Other(format!("vba.save: `kind` must be standard or class, not {other:?}"))),
    };
    let Some(bytes) = s.doc()?.wb.vba_project.clone() else {
        return Err(EngineError::Other("The workbook has no VBA project; open an .xlsm to edit its macros.".into()));
    };
    let project = Project::from_bin(&bytes).map_err(|e| EngineError::Other(format!("The VBA project can't be read: {e}")))?;
    let mut modules = modules_of(&project);
    let created = match modules.iter().position(|m| m.name.eq_ignore_ascii_case(&name)) {
        Some(i) => {
            if let Some(m) = modules.get_mut(i) {
                let header = split_attributes(&m.source).0.to_string();
                m.source = format!("{header}{code}");
            }
            false
        }
        None => {
            let header = if kind == celdas_vba::ModuleKind::Class {
                format!(
                    "Attribute VB_Name = \"{name}\"\r\nAttribute VB_GlobalNameSpace = False\r\nAttribute VB_Creatable = False\r\nAttribute VB_PredeclaredId = False\r\nAttribute VB_Exposed = False\r\n"
                )
            } else {
                format!("Attribute VB_Name = \"{name}\"\r\n")
            };
            modules.push(SourceModule { name: name.clone(), kind, source: format!("{header}{code}") });
            true
        }
    };
    let out = celdas_vba::rewrite_bin(&bytes, &modules).map_err(project_error)?;
    replace_project(s, out)?;
    let saved = Project::from_sources(vec![SourceModule { name: name.clone(), kind, source: code.to_string() }]);
    let procedures: Vec<String> = saved.procedures().into_iter().map(|p| p.name).collect();
    let errors: Vec<String> = saved.syntax_errors().into_iter().map(|(_, e)| e.to_string()).collect();
    Ok(json!({"module": name, "created": created, "procedures": procedures, "errors": errors}))
}

fn remove(s: &mut Session, p: &Json) -> Result<Json> {
    let name = str_param(p, "module")
        .map(str::trim)
        .filter(|n| !n.is_empty())
        .ok_or_else(|| EngineError::Other("vba.remove: missing `module`".into()))?
        .to_string();
    let Some(bytes) = s.doc()?.wb.vba_project.clone() else {
        return Err(EngineError::Other("The workbook has no VBA project.".into()));
    };
    let project = Project::from_bin(&bytes).map_err(|e| EngineError::Other(format!("The VBA project can't be read: {e}")))?;
    let mut modules = modules_of(&project);
    let i = modules
        .iter()
        .position(|m| m.name.eq_ignore_ascii_case(&name))
        .ok_or_else(|| EngineError::Other(format!("The VBA project has no module '{name}'.")))?;
    if modules.get(i).is_some_and(|m| is_document(&m.source)) {
        return Err(EngineError::Other(format!("'{name}' is the code behind ThisWorkbook or a sheet and can't be removed; empty its code instead.")));
    }
    modules.remove(i);
    let out = celdas_vba::rewrite_bin(&bytes, &modules).map_err(project_error)?;
    replace_project(s, out)?;
    Ok(json!({"module": name}))
}

/// `Name` from `Name`, `Module.Name`, or a shape's macro text (`'Book.xlsm'!Name`, `[0]!Name`).
fn procedure_name(text: &str) -> &str {
    let t = text.trim();
    t.rsplit_once('!').map_or(t, |(_, name)| name).trim_matches('\'')
}

fn run(s: &mut Session, p: &Json) -> Result<Json> {
    let wanted = str_param(p, "procedure").ok_or_else(|| EngineError::Other("vba.run: missing `procedure`".into()))?;
    let name = procedure_name(wanted).to_string();
    let Some((project, fingerprint)) = project_of(s)? else { return Err(EngineError::Other("The workbook has no VBA project.".into())) };
    let (doc, trust) = {
        let d = s.doc()?;
        (d.uid, d.trust)
    };
    // Only a procedure the project has (a typo must not ask for an approval first).
    let known = project.procedures().into_iter().any(|x| {
        let full = format!("{}.{}", x.module, x.name);
        x.name.eq_ignore_ascii_case(&name) || full.eq_ignore_ascii_case(&name)
    });
    if !known {
        if let Some((module, e)) = project.syntax_errors().into_iter().next() {
            return Err(EngineError::Other(format!("Compile error in {module}: {e}")));
        }
        return Err(EngineError::Other(format!("Sub or Function not defined: '{name}'")));
    }
    if trust != Trust::Local {
        let key = (doc, format!("vba:{fingerprint:x}"));
        if !s.macro_approvals.contains(&key) {
            if bool_param(p, "approve") != Some(true) {
                return Err(EngineError::Approve(format!(
                    "Run VBA macro '{name}'? It can change this workbook and read and write files on this computer."
                )));
            }
            s.macro_approvals.insert(key);
        }
    }
    let caller = str_param(p, "caller").map(str::to_string);
    let label = format!("Run VBA: {name}");
    s.grouped(&label, |s| {
        let mut host = ExcelHost::new(s);
        host.caller = caller;
        let (result, steps, site) = {
            let mut vm = Vm::new(&project, &mut host);
            let r = vm.run(&name, Vec::new());
            (r, vm.steps(), vm.error_site().cloned())
        };
        match result {
            Ok(_) => host.finish().map_err(|e| EngineError::Other(e.to_string()))?,
            Err(e) => {
                let at = site.map(|x| format!(" (in {}.{}, line {})", x.module, x.procedure, x.line)).unwrap_or_default();
                return Err(EngineError::Other(format!("{e}{at}")));
            }
        }
        let messages: Vec<Json> = host.messages.iter().map(|m| json!({"title": m.title, "text": m.text})).collect();
        Ok(json!({"procedure": name, "messages": messages, "files": host.written, "steps": steps}))
    })
}

#[cfg(test)]
mod tests;
