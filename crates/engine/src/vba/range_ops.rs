// SPDX-License-Identifier: MIT OR Apache-2.0
//! The rest of `Range` a recorded macro uses: the clipboard (`Copy`, `Cut`, `PasteSpecial`), filling,
//! sorting, filtering and hiding.

use celdas_core::col_to_letters;
use celdas_vba::{CallArg, Value, VbError, arg};
use serde_json::json;

use super::app::object_of;
use super::host::ExcelHost;
use super::objects::{RangeObj, Span, XObj};

type Result<T> = std::result::Result<T, VbError>;

/// `xlPasteAll` and the others, as `edit.pasteSpecial` names them.
fn paste_what(v: i32) -> Option<&'static str> {
    Some(match v {
        -4104 => "all",
        -4123 => "formulas",
        -4163 => "values",
        -4122 => "formats",
        -4144 => "comments",
        6 => "validation",
        7 => "allExceptBorders",
        8 => "columnWidths",
        11 => "formulasAndNumberFormats",
        12 => "valuesAndNumberFormats",
        _ => return None,
    })
}

fn paste_operation(v: i32) -> Option<&'static str> {
    Some(match v {
        -4142 => "none",
        2 => "add",
        3 => "subtract",
        4 => "multiply",
        5 => "divide",
        _ => return None,
    })
}

impl ExcelHost<'_> {
    /// The `Shift` argument of `Insert` and `Delete`, as the commands name it.
    pub(super) fn shift_of(args: &[CallArg], default: &'static str) -> Result<&'static str> {
        Ok(match arg(args, 0, "shift").map(Value::to_i32).transpose()? {
            None => default,
            Some(-4121) => "down",
            Some(-4162) => "up",
            Some(-4159) => "left",
            Some(-4161) => "right",
            Some(_) => return Err(VbError::invalid_call()),
        })
    }

    /// A range given as an argument.
    fn range_arg(args: &[CallArg], index: usize, name: &str) -> Result<Option<RangeObj>> {
        match arg(args, index, name).and_then(object_of) {
            Some(XObj::Range(r)) => Ok(Some(r)),
            _ => Ok(None),
        }
    }

    /// The members of `Range` that run a command, or `None` when `name` is not one of them.
    pub(super) fn range_action(&mut self, r: &RangeObj, name: &str, args: &[CallArg], assign: Option<Value>) -> Result<Option<Value>> {
        let done = Ok(Some(Value::Empty));
        match name {
            "filldown" => self.simple(r, "edit.fillDown").map(Some),
            "fillright" => self.simple(r, "edit.fillRight").map(Some),
            "fillup" => self.simple(r, "edit.fillUp").map(Some),
            "fillleft" => self.simple(r, "edit.fillLeft").map(Some),
            "autofilter" => self.simple(r, "data.filter").map(Some),
            "hidden" => {
                let Some(v) = assign else {
                    return Err(VbError::app("Unable to get the Hidden property of the Range class"));
                };
                let hide = v.to_bool()?;
                let (id, key) = match (r.span, hide) {
                    (Span::Rows, true) => ("home.hideRows", "rows"),
                    (Span::Rows, false) => ("home.unhideRows", "rows"),
                    (Span::Cols, true) => ("home.hideColumns", "cols"),
                    (Span::Cols, false) => ("home.unhideColumns", "cols"),
                    (Span::Cells, _) => return Err(VbError::app("Unable to set the Hidden property of the Range class")),
                };
                self.run(r.book, id, json!({"sheet": &*r.sheet, key: r.r.a1()}))?;
                done
            }
            "copy" | "cut" => {
                let dest = Self::range_arg(args, 0, "destination")?;
                match (name, dest) {
                    ("copy", Some(d)) => {
                        self.run(r.book, "edit.copy", json!({"sheet": &*r.sheet, "range": r.r.a1()}))?;
                        self.run(d.book, "edit.paste", json!({"sheet": &*d.sheet, "at": d.r.start.a1()}))?;
                        self.run(d.book, "edit.clearClipboard", json!({}))?;
                    }
                    (_, Some(d)) => {
                        self.run(r.book, "range.move", json!({"sheet": &*r.sheet, "range": r.r.a1(), "to": d.r.start.a1(), "overwrite": true}))?;
                    }
                    (_, None) => {
                        let id = if name == "copy" { "edit.copy" } else { "edit.cut" };
                        self.run(r.book, id, json!({"sheet": &*r.sheet, "range": r.r.a1()}))?;
                    }
                }
                done
            }
            "pastespecial" => {
                let what = match arg(args, 0, "paste").map(Value::to_i32).transpose()? {
                    None => "all",
                    Some(v) => paste_what(v).ok_or_else(VbError::invalid_call)?,
                };
                let operation = match arg(args, 1, "operation").map(Value::to_i32).transpose()? {
                    None => "none",
                    Some(v) => paste_operation(v).ok_or_else(VbError::invalid_call)?,
                };
                let skip = arg(args, 2, "skipblanks").map(Value::to_bool).transpose()?.unwrap_or(false);
                let transpose = arg(args, 3, "transpose").map(Value::to_bool).transpose()?.unwrap_or(false);
                // Pasted where the range is: its first cell becomes the selection.
                self.run(r.book, "sheet.activate", json!({"sheet": &*r.sheet}))?;
                self.run(r.book, "selection.set", json!({"range": r.r.a1(), "sheet": &*r.sheet}))?;
                self.run(
                    r.book,
                    "edit.pasteSpecial",
                    json!({"what": what, "operation": operation, "skipBlanks": skip, "transpose": transpose}),
                )?;
                done
            }
            "sort" => {
                let mut keys = Vec::new();
                for (key_at, order_at) in [(0, 1), (2, 4), (5, 6)] {
                    let n = keys.len() + 1;
                    let Some(k) = Self::range_arg(args, key_at, &format!("key{n}"))? else { continue };
                    let desc = arg(args, order_at, &format!("order{n}")).map(Value::to_i32).transpose()?.unwrap_or(1) == 2;
                    keys.push(json!({"column": col_to_letters(k.r.start.col), "order": if desc { "desc" } else { "asc" }}));
                }
                if keys.is_empty() {
                    return Err(VbError::invalid_call());
                }
                // xlGuess (0, the default) lets the command see whether the first row is a header.
                let header = arg(args, 7, "header").map(Value::to_i32).transpose()?.unwrap_or(0);
                let mut p = json!({"sheet": &*r.sheet, "range": r.r.a1(), "keys": keys});
                if header != 0 {
                    p["header"] = json!(header == 1);
                }
                self.run(r.book, "data.sort", p)?;
                done
            }
            _ => Ok(None),
        }
    }
}
