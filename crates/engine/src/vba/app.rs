//! `Application`, workbooks and worksheets: what a macro reaches before it gets to a range.

use std::rc::Rc;

use celdas_core::{CellRef, RangeRef};
use celdas_vba::{CallArg, Value, VbError, arg};
use serde_json::json;

use super::host::{ExcelHost, array_rows, engine_err, text_arg, to_vba};
use super::objects::{Span, X, XObj};
use crate::Session;

type Result<T> = std::result::Result<T, VbError>;

/// The part of a call's arguments that names an item of a collection: a number (1-based) or a name.
enum Item {
    Index(usize),
    Name(String),
}

fn item_arg(args: &[CallArg]) -> Result<Item> {
    match arg(args, 0, "index") {
        Some(Value::Str(s)) => Ok(Item::Name(s.to_string())),
        Some(v) => {
            let n = v.to_i32()?;
            usize::try_from(n).ok().filter(|n| *n >= 1).map(Item::Index).ok_or_else(VbError::subscript)
        }
        None => Err(VbError::invalid_call()),
    }
}

impl ExcelHost<'_> {
    pub(super) fn app_member(&mut self, name: &str, args: &[CallArg], assign: Option<Value>) -> Result<Value> {
        match name {
            "caller" => Ok(self.caller.clone().map_or(Value::Error(2023), Value::string)),
            "worksheetfunction" => Ok(X::value(XObj::WorksheetFunction)),
            "activesheet" | "activeworkbook" | "thisworkbook" | "workbooks" | "sheets" | "worksheets" | "range" | "cells" | "rows" | "columns"
            | "selection" | "activecell" => {
                use celdas_vba::Host;
                self.global(name, args, None)?.ok_or_else(|| VbError::no_member(name))
            }
            "version" => Ok(Value::str("16.0")),
            "name" => Ok(Value::str("Celdas")),
            "pathseparator" => Ok(Value::str(if cfg!(windows) { "\\" } else { "/" })),
            "username" => Ok(Value::string(self.s.prefs.user_name.clone())),
            "evaluate" | "_evaluate" => {
                let text = text_arg(args, 0, "name")?.ok_or_else(VbError::invalid_call)?;
                let book = self.active_book;
                let sheet = self.active_sheet_name(book)?;
                self.evaluate(book, &sheet, &text)
            }
            "calculate" | "calculatefull" | "doevents" | "wait" | "volatile" | "undo" => Ok(Value::Empty),
            _ => self.app_property(name, assign),
        }
    }

    /// A setting a macro turns off and on (`ScreenUpdating`…): remembered, with Excel's default when it was never set.
    fn app_property(&mut self, name: &str, assign: Option<Value>) -> Result<Value> {
        const KNOWN: &[&str] = &[
            "screenupdating",
            "displayalerts",
            "enableevents",
            "calculation",
            "statusbar",
            "cutcopymode",
            "interactive",
            "iteration",
            "displaystatusbar",
            "ignoreremoterequests",
            "asktoupdatelinks",
        ];
        if !KNOWN.contains(&name) {
            return Err(VbError::no_member(name));
        }
        if let Some(v) = assign {
            self.props.insert(name.to_string(), v);
            return Ok(Value::Empty);
        }
        Ok(self.props.get(name).cloned().unwrap_or(match name {
            "calculation" => Value::Long(-4105),
            "statusbar" | "cutcopymode" | "iteration" => Value::Bool(false),
            _ => Value::Bool(true),
        }))
    }

    /// Evaluates formula text on a sheet, as `Application.Evaluate` does.
    pub(super) fn evaluate(&mut self, book: usize, sheet: &str, text: &str) -> Result<Value> {
        self.flush()?;
        let s: &mut Session = self.session(book)?;
        let d = s.doc().map_err(engine_err)?;
        let i = d.wb.sheets.iter().position(|sh| sh.name == sheet).ok_or_else(VbError::subscript)?;
        let v = celdas_calc::evaluate(&d.wb, i, CellRef::default(), text);
        Ok(to_vba(&v, false))
    }

    /// `Application.WorksheetFunction.Name(args)`: the function evaluated on the arguments.
    pub(super) fn function_member(&mut self, name: &str, args: &[CallArg]) -> Result<Value> {
        if name == "pi" {
            return Ok(Value::Double(std::f64::consts::PI));
        }
        let mut parts = Vec::new();
        for a in args {
            parts.push(self.formula_argument(&a.value)?);
        }
        let text = format!("{}({})", name.to_uppercase(), parts.join(","));
        let book = self.active_book;
        let sheet = self.active_sheet_name(book)?;
        let v = self.evaluate(book, &sheet, &text)?;
        match v {
            Value::Error(_) => Err(VbError::app(format!("Unable to get the {name} property of the WorksheetFunction class"))),
            v => Ok(v),
        }
    }

    /// An argument written into a formula.
    fn formula_argument(&mut self, v: &Value) -> Result<String> {
        Ok(match v {
            Value::Str(s) => format!("\"{}\"", s.replace('"', "\"\"")),
            Value::Bool(b) => if *b { "TRUE" } else { "FALSE" }.into(),
            Value::Empty | Value::Missing | Value::Null => String::new(),
            Value::Object(Some(o)) => match o.downcast::<X>().map(|x| &x.0) {
                Some(XObj::Range(r)) => format!("{}!{}", celdas_core::quote_sheet(&r.sheet), r.r.a1()),
                _ => return Err(VbError::type_mismatch()),
            },
            Value::Array(_) => {
                let rows = array_rows(v).ok_or_else(VbError::type_mismatch)?;
                let mut out = Vec::new();
                for row in rows {
                    let cells: Result<Vec<String>> = row.iter().map(|c| self.formula_argument(c)).collect();
                    out.push(cells?.join(","));
                }
                format!("{{{}}}", out.join(";"))
            }
            other => celdas_vba::value::number_text(other.to_f64()?, 15),
        })
    }

    pub(super) fn workbooks_member(&mut self, name: &str, args: &[CallArg]) -> Result<Value> {
        let open: Vec<usize> = std::iter::once(0).chain(self.books.iter().enumerate().filter(|(_, s)| s.is_some()).map(|(i, _)| i + 1)).collect();
        match name {
            "count" => Ok(Value::Long(open.len() as i32)),
            "" | "item" => match item_arg(args)? {
                Item::Index(i) => open.get(i - 1).map(|b| X::value(XObj::Workbook(*b))).ok_or_else(VbError::subscript),
                Item::Name(n) => open
                    .iter()
                    .find(|b| super::host::same_name(&self.book_file_name(**b), &n))
                    .map(|b| X::value(XObj::Workbook(*b)))
                    .ok_or_else(VbError::subscript),
            },
            "add" => {
                self.flush()?;
                let mut sub = Session::new();
                sub.prefs.sheets_in_new_workbook = 1;
                sub.new_workbook();
                self.books.push(Some(sub));
                let book = self.books.len();
                self.active_book = book;
                Ok(X::value(XObj::Workbook(book)))
            }
            "open" | "close" => Err(VbError::new(70, "Permission denied: a macro can't open workbooks here")),
            _ => Err(VbError::no_member(name)),
        }
    }

    pub(super) fn shell_member(&mut self, name: &str, args: &[CallArg]) -> Result<Value> {
        match name {
            "specialfolders" | "" => {
                let which = text_arg(args, 0, "objname")?.unwrap_or_default().to_lowercase();
                let home = super::files::environ(if cfg!(windows) { "USERPROFILE" } else { "HOME" });
                let sep = if cfg!(windows) { '\\' } else { '/' };
                Ok(Value::string(match which.as_str() {
                    "desktop" | "allusersdesktop" => super::files::desktop(),
                    "mydocuments" | "documents" => format!("{home}{sep}Documents"),
                    _ => home,
                }))
            }
            "expandenvironmentstrings" => {
                let text = text_arg(args, 0, "src")?.unwrap_or_default();
                let mut out = String::new();
                let mut rest = text.as_str();
                while let Some((before, after)) = rest.split_once('%') {
                    out.push_str(before);
                    match after.split_once('%') {
                        Some((var, tail)) => {
                            out.push_str(&super::files::environ(var));
                            rest = tail;
                        }
                        None => {
                            out.push('%');
                            rest = after;
                        }
                    }
                }
                out.push_str(rest);
                Ok(Value::string(out))
            }
            "currentdirectory" => Ok(Value::string(self.book_dir(0))),
            _ => Err(VbError::new(70, "Permission denied: a macro can't run programs here")),
        }
    }

    pub(super) fn workbook_member(&mut self, book: usize, name: &str, args: &[CallArg], assign: Option<Value>) -> Result<Value> {
        match name {
            "path" => Ok(Value::string(self.book_dir(book))),
            "name" => Ok(Value::string(self.book_file_name(book))),
            "fullname" => {
                let dir = self.book_dir(book);
                let file = self.book_file_name(book);
                Ok(Value::string(if dir.is_empty() { file } else { format!("{dir}{}{file}", if cfg!(windows) { '\\' } else { '/' }) }))
            }
            "sheets" | "worksheets" => {
                if args.is_empty() {
                    Ok(X::value(XObj::Sheets(book)))
                } else {
                    self.sheets_member(book, "item", args)
                }
            }
            "activesheet" => {
                let n = self.active_sheet_name(book)?;
                Ok(X::value(XObj::Sheet(book, n)))
            }
            "application" => Ok(X::value(XObj::Application)),
            "saved" => Ok(assign.map_or(Value::Bool(false), |_| Value::Empty)),
            "activate" => {
                self.active_book = book;
                Ok(Value::Empty)
            }
            "save" => {
                self.run(book, "file.save", json!({}))?;
                Ok(Value::Empty)
            }
            "saveas" => self.save_as(book, args),
            "close" => {
                if book == 0 {
                    return Err(VbError::app("A macro can't close the workbook it runs from."));
                }
                self.flush()?;
                if let Some(slot) = self.books.get_mut(book - 1) {
                    *slot = None;
                }
                if self.active_book == book {
                    self.active_book = 0;
                }
                Ok(Value::Empty)
            }
            _ => Err(VbError::no_member(name)),
        }
    }

    fn save_as(&mut self, book: usize, args: &[CallArg]) -> Result<Value> {
        let path = text_arg(args, 0, "filename")?.ok_or_else(VbError::invalid_call)?;
        let path = super::files::norm(&path);
        let format = arg(args, 1, "fileformat").map(Value::to_i32).transpose()?;
        // A name without an extension takes the one the format has.
        let has_ext = path.rsplit(['/', '\\']).next().is_some_and(|f| f.contains('.'));
        let path = match (has_ext, format) {
            (false, Some(51)) => format!("{path}.xlsx"),
            (false, Some(52)) => format!("{path}.xlsm"),
            (false, Some(6)) => format!("{path}.csv"),
            _ => path,
        };
        self.run(book, "file.saveAs", json!({"path": path}))?;
        self.written.push(path);
        Ok(Value::Empty)
    }

    pub(super) fn sheets_member(&mut self, book: usize, name: &str, args: &[CallArg]) -> Result<Value> {
        let names = self.sheet_names(book)?;
        match name {
            "count" => Ok(Value::Long(names.len() as i32)),
            "" | "item" => {
                let n = match item_arg(args)? {
                    Item::Index(i) => names.get(i - 1).cloned().ok_or_else(VbError::subscript)?,
                    Item::Name(n) => self.canonical_sheet(book, &n)?.to_string(),
                };
                Ok(X::value(XObj::Sheet(book, Rc::from(n.as_str()))))
            }
            "add" => {
                let mut p = json!({});
                if let Some(before) = arg(args, 0, "before") {
                    if let Some(XObj::Sheet(_, n)) = object_of(before)
                        && let Some(i) = names.iter().position(|x| *x == *n)
                    {
                        p["before"] = json!(i);
                    }
                } else if let Some(after) = arg(args, 1, "after")
                    && let Some(XObj::Sheet(_, n)) = object_of(after)
                    && let Some(i) = names.iter().position(|x| *x == *n)
                {
                    p["before"] = json!(i + 1);
                }
                self.run(book, "home.insertSheet", p)?;
                // The new sheet is the active one.
                let n = self.active_sheet_name(book)?;
                Ok(X::value(XObj::Sheet(book, n)))
            }
            _ => Err(VbError::no_member(name)),
        }
    }

    pub(super) fn sheet_member(&mut self, book: usize, sheet: &Rc<str>, name: &str, args: &[CallArg], assign: Option<Value>) -> Result<Value> {
        match name {
            "name" => match assign {
                Some(v) => {
                    let to = v.to_text()?;
                    self.run(book, "sheet.rename", json!({"sheet": &**sheet, "name": to}))?;
                    Ok(Value::Empty)
                }
                None => Ok(Value::str(sheet)),
            },
            "index" => {
                let names = self.sheet_names(book)?;
                Ok(Value::Long(names.iter().position(|n| n == &**sheet).map_or(0, |i| i as i32 + 1)))
            }
            "activate" | "select" => {
                self.run(book, "sheet.activate", json!({"sheet": &**sheet}))?;
                Ok(Value::Empty)
            }
            "paste" => {
                self.run(book, "edit.paste", json!({"sheet": &**sheet}))?;
                Ok(Value::Empty)
            }
            "delete" => {
                self.run(book, "home.deleteSheet", json!({"sheet": &**sheet}))?;
                Ok(Value::Empty)
            }
            "parent" => Ok(X::value(XObj::Workbook(book))),
            "visible" => Ok(Value::Long(-1)),
            "calculate" => Ok(Value::Empty),
            "shapes" => {
                if args.is_empty() {
                    Ok(X::value(XObj::Shapes(book, sheet.clone())))
                } else {
                    self.shapes_member(book, sheet, "item", args)
                }
            }
            "evaluate" => {
                let text = text_arg(args, 0, "name")?.ok_or_else(VbError::invalid_call)?;
                self.evaluate(book, sheet, &text)
            }
            "range" => {
                let r = self.range_from_args(book, sheet, args)?;
                Ok(self.range_obj(book, sheet.clone(), r, Span::Cells))
            }
            "cells" => match (arg(args, 0, "rowindex"), arg(args, 1, "columnindex")) {
                (None, _) => Ok(self.range_obj(book, sheet.clone(), RangeRef::all(), Span::Cells)),
                (Some(r), Some(c)) => {
                    let at = cell_at(r.to_i32()?, c.to_i32()?)?;
                    Ok(self.range_obj(book, sheet.clone(), RangeRef::cell(at), Span::Cells))
                }
                // `Cells(n)`: the n-th cell, counting along the rows.
                (Some(n), None) => {
                    let n = i64::from(n.to_i32()?) - 1;
                    let at = CellRef::new(
                        u32::try_from(n / 16_384).map_err(|_| VbError::subscript())?,
                        u32::try_from(n % 16_384).map_err(|_| VbError::subscript())?,
                    );
                    Ok(self.range_obj(book, sheet.clone(), RangeRef::cell(at), Span::Cells))
                }
            },
            "rows" | "columns" => {
                let cols = name == "columns";
                let span = if cols { Span::Cols } else { Span::Rows };
                let r = match arg(args, 0, "index") {
                    None => RangeRef::all(),
                    Some(v) => line_range(v, cols)?,
                };
                Ok(self.range_obj(book, sheet.clone(), r, span))
            }
            "usedrange" => {
                self.flush()?;
                let used = self.with_sheet(book, sheet, |sh, _| sh.used_range())?;
                Ok(self.range_obj(book, sheet.clone(), used.unwrap_or_default(), Span::Cells))
            }
            "selection" | "activecell" => {
                let (sel, active) = {
                    let d = self.session(book)?.doc().map_err(engine_err)?;
                    (d.selection.current(), d.selection.active)
                };
                let r = if name == "selection" { sel } else { RangeRef::cell(active) };
                Ok(self.range_obj(book, sheet.clone(), r, Span::Cells))
            }
            _ => Err(VbError::no_member(name)),
        }
    }

    /// The cells `Range(...)` names: `"A1:B2"`, `"A1", "B2"`, or two cell objects.
    pub(super) fn range_from_args(&mut self, _book: usize, _sheet: &Rc<str>, args: &[CallArg]) -> Result<RangeRef> {
        let first = arg(args, 0, "cell1").ok_or_else(VbError::invalid_call)?;
        let one = |v: &Value| -> Result<RangeRef> {
            match v {
                Value::Object(Some(o)) => match o.downcast::<X>().map(|x| &x.0) {
                    Some(XObj::Range(r)) => Ok(r.r),
                    _ => Err(VbError::type_mismatch()),
                },
                other => parse_area(&other.to_text()?),
            }
        };
        let a = one(first)?;
        match arg(args, 1, "cell2") {
            Some(second) => Ok(a.union(&one(second)?)),
            None => Ok(a),
        }
    }
}

pub(super) fn object_of(v: &Value) -> Option<XObj> {
    match v {
        Value::Object(Some(o)) => o.downcast::<X>().map(|x| x.0.clone()),
        _ => None,
    }
}

/// `A1`, `A1:B2`, `B:B`, `3:5` (with `$`, and a sheet prefix ignored) as a range.
pub(super) fn parse_area(text: &str) -> Result<RangeRef> {
    let (_, body) = celdas_core::split_sheet(text.trim());
    if body.contains(',') {
        return Err(VbError::app("A range of several areas isn't supported here."));
    }
    RangeRef::parse(body).ok_or_else(|| VbError::app(format!("Method 'Range' of object '_Worksheet' failed: '{text}' is not a range")))
}

/// The cell at 1-based `row` and `col`.
pub(super) fn cell_at(row: i32, col: i32) -> Result<CellRef> {
    let r =
        u32::try_from(row.checked_sub(1).ok_or_else(VbError::subscript)?).map_err(|_| VbError::app("Application-defined or object-defined error"))?;
    let c =
        u32::try_from(col.checked_sub(1).ok_or_else(VbError::subscript)?).map_err(|_| VbError::app("Application-defined or object-defined error"))?;
    let at = CellRef::new(r, c);
    if at.is_valid() { Ok(at) } else { Err(VbError::app("Application-defined or object-defined error")) }
}

/// What `Rows(x)` or `Columns(x)` names: a number, `"3:5"`, `"F:R"` or `"F"`.
fn line_range(v: &Value, cols: bool) -> Result<RangeRef> {
    if let Value::Str(s) = v {
        let t = s.trim();
        if let Some(r) = RangeRef::parse(t).filter(|r| if cols { r.is_full_cols() } else { r.is_full_rows() }) {
            return Ok(r);
        }
        let single = format!("{t}:{t}");
        return RangeRef::parse(&single).ok_or_else(|| VbError::app("Application-defined or object-defined error"));
    }
    let n = u32::try_from(v.to_i32()?.checked_sub(1).ok_or_else(VbError::subscript)?).map_err(|_| VbError::subscript())?;
    let (limit, make): (u32, fn(u32, u32) -> RangeRef) =
        if cols { (celdas_core::MAX_COLS, RangeRef::cols) } else { (celdas_core::MAX_ROWS, RangeRef::rows) };
    if n >= limit {
        return Err(VbError::subscript());
    }
    Ok(make(n, n))
}
