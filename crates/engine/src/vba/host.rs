//! The Excel object model a VBA macro runs against: the workbook the macro belongs to, any
//! workbook it creates, and the dialogs and files around them.
//!
//! Every change goes through the engine's own commands, so a macro obeys sheet protection, keeps
//! formulas calculating and leaves one undo step. Cell writes wait in a buffer and go in as a few
//! blocks just before anything needs to see them: a loop that writes a table cell by cell costs one
//! command, not thousands.

use std::collections::{BTreeMap, HashMap};
use std::rc::Rc;

use celdas_core::{CellError, CellRef, RangeRef, Value as Cv};
use celdas_vba::ast::FileMode;
use celdas_vba::{CallArg, Host, Obj, Value, VbError, arg};
use serde_json::{Value as Json, json};

use super::consts::excel_constant;
use super::files::Files;
use super::objects::{RangeObj, Span, X, XObj};
use crate::{EngineError, Session};

type Result<T> = std::result::Result<T, VbError>;

/// A `MsgBox` the macro showed: the run collects them for whoever started it.
#[derive(Clone, Debug, PartialEq)]
pub struct Message {
    pub title: String,
    pub text: String,
}

/// Most cells a macro may hold in its write buffer before it goes into the workbook.
const PENDING_LIMIT: usize = 400_000;
/// Most cells one assignment may write (a whole column of values is the practical largest).
pub(super) const WRITE_LIMIT: u64 = 2_000_000;

/// Buffered cell writes of one sheet, by (row, column).
type Cells = BTreeMap<(u32, u32), Json>;
/// A run of cell rows to write: the row and its runs of (first column, values).
type Row = (u32, Vec<(u32, Vec<Json>)>);

pub struct ExcelHost<'a> {
    pub(super) s: &'a mut Session,
    /// Workbooks the macro created (`Workbooks.Add`); book `n` is `books[n - 1]`. Book 0 is the
    /// workbook the macro lives in.
    pub(super) books: Vec<Option<Session>>,
    pub(super) active_book: usize,
    /// Cells written and not yet in the workbook: (book, sheet) → (row, col) → value.
    pending: BTreeMap<(usize, Rc<str>), Cells>,
    pending_cells: usize,
    pub messages: Vec<Message>,
    /// `Application.Caller`: the shape that was clicked.
    pub caller: Option<String>,
    /// Properties of `Application` a macro sets and reads back (`ScreenUpdating`…).
    pub(super) props: HashMap<String, Value>,
    pub(super) files: Files,
    /// Charts the macro made with a smooth-line type, so the series it adds are smooth too.
    pub(super) smooth: std::collections::HashSet<(usize, Rc<str>, u32)>,
    /// Files the macro wrote, for the run's report.
    pub written: Vec<String>,
}

pub(super) fn engine_err(e: EngineError) -> VbError {
    match e {
        EngineError::Other(m) => VbError::app(m),
        other => VbError::app(other.to_string()),
    }
}

/// Sheet and workbook names compare ignoring case in every script, accents included.
pub(super) fn same_name(a: &str, b: &str) -> bool {
    a.eq_ignore_ascii_case(b) || a.chars().flat_map(char::to_lowercase).eq(b.chars().flat_map(char::to_lowercase))
}

/// The VBA number of a cell error (`CVErr`).
pub(super) fn error_number(e: CellError) -> i32 {
    match e {
        CellError::Null => 2000,
        CellError::Div0 => 2007,
        CellError::Value => 2015,
        CellError::Ref => 2023,
        CellError::Name => 2029,
        CellError::Num => 2036,
        CellError::NA => 2042,
        _ => 2015,
    }
}

fn error_from_number(n: i32) -> CellError {
    match n {
        2000 => CellError::Null,
        2007 => CellError::Div0,
        2023 => CellError::Ref,
        2029 => CellError::Name,
        2036 => CellError::Num,
        2042 => CellError::NA,
        _ => CellError::Value,
    }
}

/// A cell's value as the macro sees it. A number in a date format is a date.
pub(super) fn to_vba(v: &Cv, date: bool) -> Value {
    match v {
        Cv::Empty => Value::Empty,
        Cv::Number(n) if date => Value::Date(*n),
        Cv::Number(n) => Value::Double(*n),
        Cv::Text(t) => Value::str(t),
        Cv::Bool(b) => Value::Bool(*b),
        Cv::Error(e) => Value::Error(error_number(*e)),
        Cv::Array(a) => a.get(0, 0).map_or(Value::Empty, |x| to_vba(x, false)),
    }
}

/// What the macro assigned, as the input a cell takes. Text goes through typed input (as typing
/// does), so `"1.5"` becomes a number and `"=A1"` a formula.
pub(super) fn to_json(v: &Value) -> Result<Json> {
    Ok(match v {
        Value::Empty | Value::Null | Value::Missing => Json::Null,
        Value::Bool(b) => Json::Bool(*b),
        Value::Byte(n) => json!(*n),
        Value::Int(n) => json!(*n),
        Value::Long(n) => json!(*n),
        Value::Single(n) => number_json(f64::from(*n)),
        Value::Double(n) | Value::Date(n) => number_json(*n),
        Value::Str(s) => Json::String(s.to_string()),
        Value::Error(n) => Json::String(error_from_number(*n).as_str().to_string()),
        Value::Array(_) | Value::Object(_) => return Err(VbError::type_mismatch()),
    })
}

fn number_json(n: f64) -> Json {
    serde_json::Number::from_f64(n).map_or_else(|| Json::String("#NUM!".into()), Json::Number)
}

/// The 2-D shape of an array a macro assigned: rows of values. A one-dimensional array is a row.
pub(super) fn array_rows(v: &Value) -> Option<Vec<Vec<Value>>> {
    let Value::Array(a) = v else { return None };
    let a = a.borrow();
    match a.bounds.as_slice() {
        [(lo, hi)] => Some(vec![(*lo..=*hi).filter_map(|i| a.get(&[i]).ok()).collect()]),
        [(rl, rh), (cl, ch)] => Some((*rl..=*rh).map(|r| (*cl..=*ch).filter_map(|c| a.get(&[r, c]).ok()).collect()).collect()),
        _ => None,
    }
}

impl<'a> ExcelHost<'a> {
    pub fn new(s: &'a mut Session) -> ExcelHost<'a> {
        ExcelHost {
            s,
            books: Vec::new(),
            active_book: 0,
            pending: BTreeMap::new(),
            pending_cells: 0,
            messages: Vec::new(),
            caller: None,
            props: HashMap::new(),
            files: Files::default(),
            smooth: Default::default(),
            written: Vec::new(),
        }
    }

    pub(super) fn session(&mut self, book: usize) -> Result<&mut Session> {
        if book == 0 {
            return Ok(&mut *self.s);
        }
        self.books.get_mut(book - 1).and_then(Option::as_mut).ok_or_else(|| VbError::app("That workbook is closed."))
    }

    fn session_ref(&self, book: usize) -> Result<&Session> {
        if book == 0 {
            return Ok(&*self.s);
        }
        self.books.get(book - 1).and_then(Option::as_ref).ok_or_else(|| VbError::app("That workbook is closed."))
    }

    /// Runs engine command `id` on workbook `book` after the buffered writes went in.
    pub(super) fn run(&mut self, book: usize, id: &str, params: Json) -> Result<Json> {
        self.flush()?;
        self.session(book)?.execute(id, params).map_err(engine_err)
    }

    /// The name `name` has in workbook `book` (found ignoring case), or an error that says there is no such sheet.
    pub(super) fn canonical_sheet(&self, book: usize, name: &str) -> Result<Rc<str>> {
        let d = self.session_ref(book)?.doc().map_err(engine_err)?;
        d.wb.sheets
            .iter()
            .find(|sh| same_name(&sh.name, name))
            .map(|sh| Rc::from(sh.name.as_str()))
            .ok_or_else(|| VbError::new(9, "Subscript out of range"))
    }

    pub(super) fn sheet_names(&self, book: usize) -> Result<Vec<String>> {
        Ok(self.session_ref(book)?.doc().map_err(engine_err)?.wb.sheets.iter().map(|s| s.name.clone()).collect())
    }

    pub(super) fn active_sheet_name(&self, book: usize) -> Result<Rc<str>> {
        let d = self.session_ref(book)?.doc().map_err(engine_err)?;
        d.wb.sheet(d.wb.active_sheet()).map(|s| Rc::from(s.name.as_str())).ok_or_else(|| VbError::app("The workbook has no sheets."))
    }

    /// Reads from sheet `sheet` of `book`: `f` sees the sheet as it is now.
    pub(super) fn with_sheet<R>(
        &mut self,
        book: usize,
        sheet: &str,
        f: impl FnOnce(&celdas_model::Sheet, &celdas_model::Workbook) -> R,
    ) -> Result<R> {
        let d = self.session_ref(book)?.doc().map_err(engine_err)?;
        let i = d.wb.sheets.iter().position(|sh| sh.name == sheet).ok_or_else(|| VbError::new(9, "Subscript out of range"))?;
        let sh = d.wb.sheet(i).ok_or_else(VbError::subscript)?;
        Ok(f(sh, &d.wb))
    }

    // ---- the write buffer

    pub(super) fn queue_write(&mut self, book: usize, sheet: &Rc<str>, at: CellRef, v: Json) -> Result<()> {
        let rows = self.pending.entry((book, sheet.clone())).or_default();
        if rows.insert((at.row, at.col), v).is_none() {
            self.pending_cells += 1;
        }
        if self.pending_cells > PENDING_LIMIT {
            self.flush()?;
        }
        Ok(())
    }

    /// The value waiting for a cell, when it is one a read can answer without the workbook.
    pub(super) fn pending_value(&self, book: usize, sheet: &Rc<str>, at: CellRef) -> Option<Value> {
        let v = self.pending.get(&(book, sheet.clone()))?.get(&(at.row, at.col))?;
        match v {
            Json::Null => Some(Value::Empty),
            Json::Bool(b) => Some(Value::Bool(*b)),
            Json::Number(n) => n.as_f64().map(Value::Double),
            _ => None,
        }
    }

    /// Puts the buffered writes into their workbooks as blocks of rows with the same columns.
    pub(super) fn flush(&mut self) -> Result<()> {
        if self.pending.is_empty() {
            return Ok(());
        }
        let pending = std::mem::take(&mut self.pending);
        self.pending_cells = 0;
        for ((book, sheet), cells) in pending {
            for (start, rows) in blocks(&cells) {
                let params = json!({"sheet": &*sheet, "range": start.a1(), "values": rows});
                self.session(book)?.execute("range.setValues", params).map_err(engine_err)?;
            }
        }
        Ok(())
    }

    /// Closes everything the macro opened and puts the buffered writes in.
    pub fn finish(&mut self) -> Result<()> {
        self.flush()?;
        self.files.close_all(&mut self.written).map_err(|e| VbError::new(75, e))
    }

    pub(super) fn range_obj(&self, book: usize, sheet: Rc<str>, r: RangeRef, span: Span) -> Value {
        X::value(XObj::Range(RangeObj { book, sheet, r, span }))
    }

    /// The directory the workbook was opened from (`ThisWorkbook.Path`).
    pub(super) fn book_dir(&self, book: usize) -> String {
        let Some(path) = self.session_ref(book).ok().and_then(|s| s.doc().ok()).and_then(|d| d.path.clone()) else { return String::new() };
        let norm = path.replace('\\', "/");
        let dir = norm.rsplit_once('/').map_or("", |(d, _)| d);
        if cfg!(windows) { dir.replace('/', "\\") } else { dir.to_string() }
    }

    pub(super) fn book_file_name(&self, book: usize) -> String {
        let Ok(d) = self.session_ref(book).and_then(|s| s.doc().map_err(engine_err)) else { return String::new() };
        match &d.path {
            Some(p) => p.replace('\\', "/").rsplit('/').next().unwrap_or("").to_string(),
            None => d.display_title(),
        }
    }
}

/// Rectangles of the cells that have a value: consecutive rows with the same run of columns join.
fn blocks(cells: &BTreeMap<(u32, u32), Json>) -> Vec<(CellRef, Vec<Vec<Json>>)> {
    // Each row's runs of adjacent columns, as (first column, values).
    let mut rows: Vec<Row> = Vec::new();
    for (&(row, col), v) in cells {
        if rows.last().is_none_or(|(r, _)| *r != row) {
            rows.push((row, Vec::new()));
        }
        let Some((_, runs)) = rows.last_mut() else { continue };
        match runs.last_mut() {
            Some((first, vals)) if first.saturating_add(vals.len() as u32) == col => vals.push(v.clone()),
            _ => runs.push((col, vec![v.clone()])),
        }
    }
    let mut out: Vec<(CellRef, Vec<Vec<Json>>, u32, usize)> = Vec::new();
    for (row, runs) in rows {
        for (col, vals) in runs {
            // Joins the block just above when its last row is right above this one and it spans these columns.
            let joined = out.last_mut().filter(|(at, rows, c, w)| *c == col && *w == vals.len() && at.row.saturating_add(rows.len() as u32) == row);
            match joined {
                Some((_, rows, ..)) => rows.push(vals),
                None => {
                    let width = vals.len();
                    out.push((CellRef::new(row, col), vec![vals], col, width));
                }
            }
        }
    }
    out.into_iter().map(|(at, rows, ..)| (at, rows)).collect()
}

impl Host for ExcelHost<'_> {
    fn member(&mut self, obj: &Obj, name: &str, args: &[CallArg], assign: Option<Value>) -> Result<Value> {
        let Some(x) = obj.downcast::<X>() else { return Err(VbError::no_member(name)) };
        let x = x.0.clone();
        // A range assigned to a property stands for its value, except where the property takes the range itself.
        let takes_range = matches!(&x, XObj::Series(..)) && matches!(name, "xvalues" | "values" | "name");
        let assign = match assign {
            Some(Value::Object(Some(o))) if !takes_range => match o.downcast::<X>().map(|v| v.0.clone()) {
                Some(XObj::Range(src)) => Some(self.range_member(&src, "value", &[], None)?),
                _ => Some(Value::Object(Some(o))),
            },
            other => other,
        };
        match x {
            XObj::Application => self.app_member(name, args, assign),
            XObj::WorksheetFunction => self.function_member(name, args),
            XObj::Workbooks => self.workbooks_member(name, args),
            XObj::Shell => self.shell_member(name, args),
            XObj::Workbook(b) => self.workbook_member(b, name, args, assign),
            XObj::Sheets(b) => self.sheets_member(b, name, args),
            XObj::Sheet(b, n) => self.sheet_member(b, &n, name, args, assign),
            XObj::Range(r) => self.range_member(&r, name, args, assign),
            XObj::Font(r) => self.font_member(&r, name, assign),
            XObj::Interior(r) => self.interior_member(&r, name, assign),
            XObj::Borders(r, i) => self.borders_member(&r, i, name, args, assign),
            XObj::Shapes(b, n) => self.shapes_member(b, &n, name, args),
            XObj::Shape(k) => self.shape_member(&k, name, args, assign),
            XObj::TextFrame(k) => self.text_frame_member(&k, name, args, assign),
            XObj::TextRange(k) => self.text_range_member(&k, name, args, assign),
            XObj::Chart(k) => self.chart_member(&k, name, args, assign),
            XObj::SeriesCollection(k) => self.series_collection_member(&k, name, args),
            XObj::Series(k, i) => self.series_member(&k, i, name, assign),
            XObj::ChartTitle(k) => self.chart_title_member(&k, name, assign),
        }
    }

    fn global(&mut self, name: &str, args: &[CallArg], assign: Option<Value>) -> Result<Option<Value>> {
        if args.is_empty()
            && assign.is_none()
            && let Some(c) = excel_constant(name)
        {
            return Ok(Some(Value::Long(c)));
        }
        let book = self.active_book;
        let v = match name {
            "application" => X::value(XObj::Application),
            "thisworkbook" => X::value(XObj::Workbook(0)),
            "activeworkbook" => X::value(XObj::Workbook(book)),
            "workbooks" if args.is_empty() => X::value(XObj::Workbooks),
            "workbooks" => self.workbooks_member("item", args)?,
            "activesheet" => {
                let n = self.active_sheet_name(book)?;
                X::value(XObj::Sheet(book, n))
            }
            "sheets" | "worksheets" if args.is_empty() => X::value(XObj::Sheets(book)),
            "sheets" | "worksheets" => self.sheets_member(book, "item", args)?,
            "range" | "cells" | "rows" | "columns" | "selection" | "activecell" | "usedrange" => {
                let n = self.active_sheet_name(book)?;
                return self.sheet_member(book, &n, name, args, assign).map(Some);
            }
            _ => return Ok(None),
        };
        if assign.is_some() {
            return Err(VbError::new(
                451,
                format!("Property let procedure not defined and property get procedure did not return an object: '{name}'"),
            ));
        }
        Ok(Some(v))
    }

    fn enumerate(&mut self, obj: &Obj) -> Result<Vec<Value>> {
        let Some(x) = obj.downcast::<X>() else { return Err(VbError::new(438, "This object cannot be enumerated")) };
        match x.0.clone() {
            XObj::Range(r) => self.range_cells(&r),
            XObj::Sheets(b) => self.sheet_names(b)?.into_iter().map(|n| Ok(X::value(XObj::Sheet(b, Rc::from(n.as_str()))))).collect(),
            XObj::Shapes(b, n) => self.shape_list(b, &n),
            XObj::SeriesCollection(k) => self.series_list(&k),
            XObj::Workbooks => Ok(std::iter::once(0)
                .chain(self.books.iter().enumerate().filter(|(_, s)| s.is_some()).map(|(i, _)| i + 1))
                .map(|b| X::value(XObj::Workbook(b)))
                .collect()),
            _ => Err(VbError::new(438, "This object cannot be enumerated")),
        }
    }

    fn create_object(&mut self, prog_id: &str) -> Result<Value> {
        match prog_id {
            "wscript.shell" => Ok(X::value(XObj::Shell)),
            other => Err(VbError::new(429, format!("ActiveX component can't create object: '{other}'"))),
        }
    }

    fn msg_box(&mut self, prompt: &str, buttons: i32, title: &str) -> i32 {
        self.messages.push(Message { title: title.to_string(), text: prompt.to_string() });
        // A question answers with its default button; a macro cannot wait for a person here.
        match buttons & 0xF {
            1 => 1,
            2 => 3,
            3 | 4 => 6,
            5 => 4,
            _ => 1,
        }
    }

    fn input_box(&mut self, _prompt: &str, _title: &str, default: &str) -> Option<String> {
        Some(default.to_string())
    }

    fn file_open(&mut self, path: &str, mode: FileMode, number: i32) -> Result<()> {
        self.files.open(path, mode, number).map_err(|e| VbError::new(76, e))
    }

    fn file_write(&mut self, number: i32, text: &str) -> Result<()> {
        self.files.write(number, text).map_err(|e| VbError::new(52, e))
    }

    fn file_read_line(&mut self, number: i32) -> Result<Option<String>> {
        self.files.read_line(number).map_err(|e| VbError::new(52, e))
    }

    fn file_eof(&mut self, number: i32) -> Result<bool> {
        self.files.at_eof(number).map_err(|e| VbError::new(52, e))
    }

    fn file_close(&mut self, number: Option<i32>) -> Result<()> {
        let r = match number {
            Some(n) => self.files.close(n, &mut self.written),
            None => self.files.close_all(&mut self.written),
        };
        r.map_err(|e| VbError::new(52, e))
    }

    fn dir(&mut self, pattern: &str, attributes: i32) -> Result<String> {
        self.files.dir(pattern, attributes).map_err(|e| VbError::new(52, e))
    }

    fn make_dir(&mut self, path: &str) -> Result<()> {
        super::files::make_dir(path).map_err(|e| VbError::new(75, e))
    }

    fn remove_file(&mut self, path: &str) -> Result<()> {
        super::files::remove_file(path).map_err(|e| VbError::new(53, e))
    }

    fn environ(&mut self, name: &str) -> String {
        super::files::environ(name)
    }
}

/// An optional argument as text.
pub(super) fn text_arg(args: &[CallArg], i: usize, name: &str) -> Result<Option<String>> {
    arg(args, i, name).map(Value::to_text).transpose()
}
