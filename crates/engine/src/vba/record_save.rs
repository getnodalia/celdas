// SPDX-License-Identifier: MIT OR Apache-2.0
//! Saving a recording as Visual Basic: the recorded commands become a `Sub` in a new module of the
//! workbook's VBA project (the project is made when the workbook has none).

use celdas_model::StepKind;
use serde_json::{Value as Json, json};

use super::record::{procedure_name, recorded_sub};
use crate::cmd::str_param;
use crate::macros::recorder::RecordingStart;
use crate::record::Recorder;
use crate::{EngineError, Result, Session};

/// A procedure name the project does not use yet (compared without regard to case): `name`, or
/// `name_2`, `name_3`…
fn free_name(taken: &[String], name: &str) -> String {
    let used = |n: &str| taken.iter().any(|t| t.eq_ignore_ascii_case(n));
    if !used(name) {
        return name.to_string();
    }
    (2..).map(|i| format!("{name}_{i}")).find(|n| !used(n)).unwrap_or_else(|| name.to_string())
}

/// The first `ModuleN` the project does not have.
fn next_module(modules: &[String]) -> String {
    (1..).map(|n| format!("Module{n}")).find(|m| !modules.iter().any(|x| x.eq_ignore_ascii_case(m))).unwrap_or_else(|| "Module1".into())
}

/// Saves `rec` into the workbook at index `at` as a procedure, and says what it did.
pub(crate) fn save(s: &mut Session, p: &Json, rec: &Recorder, start: &RecordingStart, at: usize) -> Result<Json> {
    let d = s.docs.get(at).ok_or(EngineError::NoDocument)?;
    let wb = d.wb.clone();
    let wanted = str_param(p, "name").map(str::trim).filter(|n| !n.is_empty()).or(start.name.as_deref());
    let base = match wanted {
        Some(n) => procedure_name(n),
        None => "Macro1".into(),
    };
    // The recording as plain commands: the same expansion as a macro of steps, in plain references.
    let converted = crate::macros::recorder::convert(&base, start, &rec.steps, &rec.ran);
    let steps: Vec<(String, Json)> = converted
        .mac
        .steps
        .iter()
        .filter_map(|st| match &st.kind {
            StepKind::Command { id, params } => Some((id.clone(), params.clone())),
            _ => None,
        })
        .collect();
    let existing = if wb.vba_project.is_some() {
        let active = s.active_index();
        s.set_active(at);
        let listed = s.execute("vba.list", json!({}));
        s.set_active(active);
        Some(listed?)
    } else {
        None
    };
    let names = |key: &str, field: &str| -> Vec<String> {
        existing
            .as_ref()
            .and_then(|l| l.get(key))
            .and_then(Json::as_array)
            .map(|a| a.iter().filter_map(|m| m.get(field).and_then(Json::as_str).map(str::to_string)).collect())
            .unwrap_or_default()
    };
    let name = match wanted {
        Some(_) => free_name(&names("procedures", "name"), &base),
        None => (1..)
            .map(|n| format!("Macro{n}"))
            .find(|n| !names("procedures", "name").iter().any(|t| t.eq_ignore_ascii_case(n)))
            .unwrap_or(base),
    };
    let description = str_param(p, "description").or(start.description.as_deref()).unwrap_or("");
    let (sub, skipped) = recorded_sub(&name, description, &steps);
    let module = if existing.is_some() { next_module(&names("modules", "name")) } else { "Module1".into() };
    let active = s.active_index();
    s.set_active(at);
    let saved = if existing.is_some() {
        s.execute("vba.save", json!({"module": module, "code": sub, "kind": "standard"}))
    } else {
        s.execute("vba.create", json!({"code": sub}))
    };
    s.set_active(active);
    saved?;
    let dropped = rec.dropped.saturating_add(converted.cut);
    let mut out = json!({
        "recording": false,
        "as": "vba",
        "saved": name,
        "module": module,
        "steps": steps.len().saturating_sub(skipped.len()),
        "dropped": dropped,
        "skipped": skipped,
        "inputs": [],
    });
    // A procedure has no shortcut of its own here.
    if let Some(k) = &start.shortcut {
        out["shortcutDropped"] = json!(k);
    }
    Ok(out)
}
