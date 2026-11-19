//! `Range`: cells, their values and formulas, formatting, and the structural calls on them.

use celdas_core::{CellRef, MAX_COLS, MAX_ROWS, RangeRef, Value as Cv, col_to_letters};
use celdas_model::{BorderLine, BorderStyle, Color, Fill, HAlign, Style, Underline, VAlign};
use celdas_vba::{CallArg, VArray, Value, VbError, arg};
use serde_json::{Value as Json, json};

use super::host::{ExcelHost, WRITE_LIMIT, array_rows, engine_err, to_json, to_vba};
use super::objects::{RangeObj, Span, X, XObj};

type Result<T> = std::result::Result<T, VbError>;

/// Most cells one read may return as an array.
const READ_LIMIT: u64 = 5_000_000;

fn failed(what: &str) -> VbError {
    VbError::app(format!("{what} method of Range class failed"))
}

/// A colour as VBA holds it (`RGB(r, g, b)`: blue in the high byte).
fn color_of(v: i64) -> Color {
    Color::rgb((v & 0xFF) as u8, ((v >> 8) & 0xFF) as u8, ((v >> 16) & 0xFF) as u8)
}

fn vba_color(c: Color) -> Value {
    match c {
        Color::Rgb(v) => Value::Long((((v & 0xFF) << 16) | (v & 0xFF00) | ((v >> 16) & 0xFF)) as i32),
        _ => Value::Long(0),
    }
}

fn h_align(v: i32) -> Result<HAlign> {
    Ok(match v {
        1 => HAlign::General,
        -4131 => HAlign::Left,
        -4108 => HAlign::Center,
        -4152 => HAlign::Right,
        5 => HAlign::Fill,
        -4130 => HAlign::Justify,
        7 => HAlign::CenterAcross,
        -4117 => HAlign::Distributed,
        _ => return Err(VbError::app("Unable to set the HorizontalAlignment property of the Range class")),
    })
}

fn h_align_value(a: HAlign) -> i32 {
    match a {
        HAlign::General => 1,
        HAlign::Left => -4131,
        HAlign::Center => -4108,
        HAlign::Right => -4152,
        HAlign::Fill => 5,
        HAlign::Justify => -4130,
        HAlign::CenterAcross => 7,
        HAlign::Distributed => -4117,
    }
}

fn v_align(v: i32) -> Result<VAlign> {
    Ok(match v {
        -4160 => VAlign::Top,
        -4108 => VAlign::Center,
        -4107 => VAlign::Bottom,
        -4130 => VAlign::Justify,
        -4117 => VAlign::Distributed,
        _ => return Err(VbError::app("Unable to set the VerticalAlignment property of the Range class")),
    })
}

fn v_align_value(a: VAlign) -> i32 {
    match a {
        VAlign::Top => -4160,
        VAlign::Center => -4108,
        VAlign::Bottom => -4107,
        VAlign::Justify => -4130,
        VAlign::Distributed => -4117,
    }
}

/// The border style a line style and weight make.
fn border_style(line: i32, weight: Option<i32>) -> BorderStyle {
    match line {
        -4142 => BorderStyle::None,
        -4115 => {
            if weight == Some(-4138) {
                BorderStyle::MediumDashed
            } else {
                BorderStyle::Dashed
            }
        }
        -4118 => BorderStyle::Dotted,
        -4119 => BorderStyle::Double,
        4 => {
            if weight == Some(-4138) {
                BorderStyle::MediumDashDot
            } else {
                BorderStyle::DashDot
            }
        }
        5 => BorderStyle::DashDotDot,
        13 => BorderStyle::SlantDashDot,
        _ => match weight {
            Some(1) => BorderStyle::Hair,
            Some(-4138) => BorderStyle::Medium,
            Some(4) => BorderStyle::Thick,
            _ => BorderStyle::Thin,
        },
    }
}

fn line_style_value(s: BorderStyle) -> i32 {
    match s {
        BorderStyle::None => -4142,
        BorderStyle::Dashed | BorderStyle::MediumDashed => -4115,
        BorderStyle::Dotted | BorderStyle::Hair => -4118,
        BorderStyle::Double => -4119,
        BorderStyle::DashDot | BorderStyle::MediumDashDot => 4,
        BorderStyle::DashDotDot => 5,
        BorderStyle::SlantDashDot => 13,
        _ => 1,
    }
}

fn weight_value(s: BorderStyle) -> i32 {
    match s {
        BorderStyle::Hair => 1,
        BorderStyle::Medium | BorderStyle::MediumDashed | BorderStyle::MediumDashDot => -4138,
        BorderStyle::Thick | BorderStyle::Double => 4,
        _ => 2,
    }
}

/// Which side of a cell a border index names.
#[derive(Clone, Copy, PartialEq, Eq)]
enum Side {
    Left,
    Right,
    Top,
    Bottom,
    DiagDown,
    DiagUp,
}

fn line_mut(b: &mut celdas_model::Borders, side: Side) -> &mut BorderLine {
    match side {
        Side::Left => &mut b.left,
        Side::Right => &mut b.right,
        Side::Top => &mut b.top,
        Side::Bottom => &mut b.bottom,
        Side::DiagDown => &mut b.diag_down,
        Side::DiagUp => &mut b.diag_up,
    }
}

/// A border line to put on `side`: what changes of the line that is there.
#[derive(Clone, Copy)]
enum Edit {
    Line(i32),
    Weight(i32),
    Color(Color),
}

fn edited(old: BorderLine, e: Edit) -> BorderLine {
    match e {
        Edit::Line(l) => BorderLine { style: border_style(l, (old.style != BorderStyle::None).then_some(weight_value(old.style))), color: old.color },
        Edit::Weight(w) => BorderLine { style: border_style(line_style_value(old.style), Some(w)), color: old.color },
        Edit::Color(c) => BorderLine { style: if old.style == BorderStyle::None { BorderStyle::Thin } else { old.style }, color: c },
    }
}

impl ExcelHost<'_> {
    fn range_a1(r: &RangeObj) -> serde_json::Value {
        json!({"sheet": &*r.sheet, "range": r.r.a1()})
    }

    /// Applies `f` to the style of the cells of `r`.
    fn restyle(&mut self, r: &RangeObj, f: impl Fn(&mut Style)) -> Result<()> {
        self.flush()?;
        let p = Self::range_a1(r);
        let s = self.session(r.book)?;
        crate::cmd::format::apply_style(s, &p, f).map_err(engine_err)?;
        Ok(())
    }

    /// The style of the first cell of `r`.
    fn first_style(&mut self, r: &RangeObj) -> Result<Style> {
        self.flush()?;
        let at = r.r.start;
        self.with_sheet(r.book, &r.sheet, |sh, wb| wb.styles.get(sh.style_id(at)).clone())
    }

    fn number_of_cells(r: &RangeObj) -> u64 {
        r.r.count()
    }

    /// The cell at 1-based (`row`, `col`) counted from the top left of `r`.
    fn relative_cell(r: &RangeObj, row: i32, col: i32) -> Result<CellRef> {
        let dr = i64::from(row) - 1;
        let dc = i64::from(col) - 1;
        r.r.start.offset(dr, dc).filter(|c| c.is_valid()).ok_or_else(|| VbError::app("Application-defined or object-defined error"))
    }

    fn sub_range(&self, r: &RangeObj, area: RangeRef, span: Span) -> Value {
        self.range_obj(r.book, r.sheet.clone(), area, span)
    }

    pub(super) fn range_member(&mut self, r: &RangeObj, name: &str, args: &[CallArg], assign: Option<Value>) -> Result<Value> {
        match name {
            "" | "value" | "value2" => match assign {
                Some(v) => self.set_values(r, &v).map(|()| Value::Empty),
                None => self.read_values(r, args),
            },
            "formula" | "formular1c1" | "formulalocal" | "formulaarray" => match assign {
                Some(v) => self.set_formula(r, &v).map(|()| Value::Empty),
                None => {
                    self.flush()?;
                    let at = r.r.start;
                    let t = self.with_sheet(r.book, &r.sheet, |sh, wb| crate::display::edit_text(wb, sh, at))?;
                    Ok(Value::string(t))
                }
            },
            "text" => {
                self.flush()?;
                let at = r.r.start;
                let t = self.with_sheet(r.book, &r.sheet, |sh, wb| crate::display::cell_text(wb, sh, at))?;
                Ok(Value::string(t))
            }
            "hasformula" => {
                self.flush()?;
                let at = r.r.start;
                Ok(Value::Bool(self.with_sheet(r.book, &r.sheet, |sh, _| sh.cell(at).is_some_and(|c| c.formula.is_some()))?))
            }
            "numberformat" | "numberformatlocal" => match assign {
                Some(v) => {
                    let code = v.to_text()?;
                    self.run(r.book, "home.numberFormat", json!({"sheet": &*r.sheet, "range": r.r.a1(), "code": code}))?;
                    Ok(Value::Empty)
                }
                None => Ok(Value::string(self.first_style(r)?.num_fmt.as_str().to_string())),
            },
            "count" => {
                let n = match r.span {
                    Span::Rows => u64::from(r.r.height()),
                    Span::Cols => u64::from(r.r.width()),
                    Span::Cells => Self::number_of_cells(r),
                };
                Ok(i32::try_from(n).map_or(Value::Double(n as f64), Value::Long))
            }
            "row" => Ok(Value::Long(r.r.start.row as i32 + 1)),
            "column" => Ok(Value::Long(r.r.start.col as i32 + 1)),
            "address" => Ok(Value::string(address(&r.r, args))),
            "left" | "top" | "width" | "height" => {
                let rr = r.r;
                let v = self.with_sheet(r.book, &r.sheet, |sh, _| match name {
                    "left" => sh.col_left(rr.start.col),
                    "top" => sh.row_top(rr.start.row),
                    "width" => sh.col_left(rr.end.col.saturating_add(1)) - sh.col_left(rr.start.col),
                    _ => sh.row_top(rr.end.row.saturating_add(1)) - sh.row_top(rr.start.row),
                })?;
                Ok(Value::Double(v))
            }
            "cells" => match (arg(args, 0, "rowindex"), arg(args, 1, "columnindex")) {
                (None, _) => Ok(self.sub_range(r, r.r, Span::Cells)),
                (Some(a), Some(b)) => {
                    let at = Self::relative_cell(r, a.to_i32()?, b.to_i32()?)?;
                    Ok(self.sub_range(r, RangeRef::cell(at), Span::Cells))
                }
                (Some(n), None) => {
                    let n = i64::from(n.to_i32()?) - 1;
                    let w = i64::from(r.r.width());
                    let at = r.r.start.offset(n / w.max(1), n % w.max(1)).filter(|c| c.is_valid()).ok_or_else(VbError::subscript)?;
                    Ok(self.sub_range(r, RangeRef::cell(at), Span::Cells))
                }
            },
            "rows" | "columns" | "entirerow" | "entirecolumn" => {
                let cols = name == "columns" || name == "entirecolumn";
                let span = if cols { Span::Cols } else { Span::Rows };
                let whole = name.starts_with("entire");
                match arg(args, 0, "index") {
                    _ if whole => {
                        let area = if cols { RangeRef::cols(r.r.start.col, r.r.end.col) } else { RangeRef::rows(r.r.start.row, r.r.end.row) };
                        Ok(self.sub_range(r, area, span))
                    }
                    None => Ok(self.sub_range(r, r.r, span)),
                    Some(Value::Str(s)) if !cols || r.r.is_full_cols() || r.r.start.col == 0 => {
                        let t = s.trim();
                        let single = format!("{t}:{t}");
                        let area =
                            RangeRef::parse(t).filter(|x| x.is_full_cols()).or_else(|| RangeRef::parse(&single)).ok_or_else(|| failed("Columns"))?;
                        Ok(self.sub_range(r, area, span))
                    }
                    Some(v) => {
                        let i = i64::from(v.to_i32()?) - 1;
                        let area = if cols {
                            let c = i64::from(r.r.start.col) + i;
                            let c = u32::try_from(c).ok().filter(|c| *c <= r.r.end.col.max(MAX_COLS - 1)).ok_or_else(VbError::subscript)?;
                            RangeRef::new(CellRef::new(r.r.start.row, c), CellRef::new(r.r.end.row, c))
                        } else {
                            let w = i64::from(r.r.start.row) + i;
                            let w = u32::try_from(w).ok().filter(|w| *w < MAX_ROWS).ok_or_else(VbError::subscript)?;
                            RangeRef::new(CellRef::new(w, r.r.start.col), CellRef::new(w, r.r.end.col))
                        };
                        Ok(self.sub_range(r, area, span))
                    }
                }
            }
            "offset" => {
                let dr = arg(args, 0, "rowoffset").map(Value::to_i32).transpose()?.unwrap_or(0);
                let dc = arg(args, 1, "columnoffset").map(Value::to_i32).transpose()?.unwrap_or(0);
                let a = r.r.start.offset(i64::from(dr), i64::from(dc)).filter(|c| c.is_valid());
                let b = r.r.end.offset(i64::from(dr), i64::from(dc)).filter(|c| c.is_valid());
                match (a, b) {
                    (Some(a), Some(b)) => Ok(self.sub_range(r, RangeRef::new(a, b), r.span)),
                    _ => Err(VbError::app("Application-defined or object-defined error")),
                }
            }
            "resize" => {
                let rows = arg(args, 0, "rowsize").map(Value::to_i32).transpose()?.unwrap_or(r.r.height() as i32);
                let cols = arg(args, 1, "columnsize").map(Value::to_i32).transpose()?.unwrap_or(r.r.width() as i32);
                if rows < 1 || cols < 1 {
                    return Err(VbError::invalid_call());
                }
                let end = r.r.start.offset(i64::from(rows) - 1, i64::from(cols) - 1).filter(|c| c.is_valid()).ok_or_else(VbError::invalid_call)?;
                Ok(self.sub_range(r, RangeRef::new(r.r.start, end), Span::Cells))
            }
            "range" => {
                let inner = self.range_from_args(r.book, &r.sheet, args)?;
                let origin = r.r.start;
                let a = origin.offset(i64::from(inner.start.row), i64::from(inner.start.col));
                let b = origin.offset(i64::from(inner.end.row), i64::from(inner.end.col));
                match (a, b) {
                    (Some(a), Some(b)) if a.is_valid() && b.is_valid() => Ok(self.sub_range(r, RangeRef::new(a, b), Span::Cells)),
                    _ => Err(failed("Range")),
                }
            }
            "end" => {
                let dir = arg(args, 0, "direction").ok_or_else(VbError::invalid_call)?.to_i32()?;
                let (dr, dc) = match dir {
                    -4162 => (-1, 0),
                    -4121 => (1, 0),
                    -4159 => (0, -1),
                    -4161 => (0, 1),
                    _ => return Err(VbError::invalid_call()),
                };
                self.flush()?;
                let start = r.r.start;
                let at = self.with_sheet(r.book, &r.sheet, |sh, _| end_of(sh, start, dr, dc))?;
                Ok(self.sub_range(r, RangeRef::cell(at), Span::Cells))
            }
            "currentregion" => {
                self.flush()?;
                let start = r.r.start;
                let area = self.with_sheet(r.book, &r.sheet, |sh, _| current_region(sh, start))?;
                Ok(self.sub_range(r, area, Span::Cells))
            }
            "clearcontents" => self.simple(r, "edit.clearContents"),
            "clearformats" => self.simple(r, "edit.clearFormats"),
            "clear" => self.simple(r, "edit.clearAll"),
            "clearcomments" => self.simple(r, "edit.clearComments"),
            "merge" => {
                let across = arg(args, 0, "across").map(Value::to_bool).transpose()?.unwrap_or(false);
                self.simple(r, if across { "home.mergeAcross" } else { "home.mergeCells" })
            }
            "unmerge" => self.simple(r, "home.unmergeCells"),
            "mergecells" => {
                self.flush()?;
                let at = r.r.start;
                Ok(Value::Bool(self.with_sheet(r.book, &r.sheet, |sh, _| sh.merge_at(at).is_some())?))
            }
            "autofit" => match r.span {
                Span::Cols => self.run_value(r.book, "home.autofitColumnWidth", json!({"sheet": &*r.sheet, "cols": r.r.a1()})),
                Span::Rows => self.run_value(r.book, "home.autofitRowHeight", json!({"sheet": &*r.sheet, "rows": r.r.a1()})),
                Span::Cells => Err(failed("AutoFit")),
            },
            "rowheight" => match assign {
                Some(v) => {
                    let h = v.to_f64()?;
                    let rows = RangeRef::rows(r.r.start.row, r.r.end.row).a1();
                    self.run_value(r.book, "home.rowHeight", json!({"sheet": &*r.sheet, "rows": rows, "points": h}))
                }
                None => {
                    let row = r.r.start.row;
                    Ok(Value::Double(f64::from(self.with_sheet(r.book, &r.sheet, |sh, _| sh.row_height(row))?)))
                }
            },
            "columnwidth" => match assign {
                Some(v) => {
                    let w = v.to_f64()?;
                    let cols = RangeRef::cols(r.r.start.col, r.r.end.col).a1();
                    self.run_value(r.book, "home.columnWidth", json!({"sheet": &*r.sheet, "cols": cols, "chars": w}))
                }
                None => {
                    let col = r.r.start.col;
                    let pts = self.with_sheet(r.book, &r.sheet, |sh, _| sh.col_width(col))?;
                    Ok(Value::Double(celdas_model::points_to_chars(f64::from(pts))))
                }
            },
            "horizontalalignment" => match assign {
                Some(v) => {
                    let a = h_align(v.to_i32()?)?;
                    self.restyle(r, move |st| st.align.h = a).map(|()| Value::Empty)
                }
                None => Ok(Value::Long(h_align_value(self.first_style(r)?.align.h))),
            },
            "verticalalignment" => match assign {
                Some(v) => {
                    let a = v_align(v.to_i32()?)?;
                    self.restyle(r, move |st| st.align.v = a).map(|()| Value::Empty)
                }
                None => Ok(Value::Long(v_align_value(self.first_style(r)?.align.v))),
            },
            "wraptext" => match assign {
                Some(v) => {
                    let on = v.to_bool()?;
                    self.restyle(r, move |st| st.align.wrap = on).map(|()| Value::Empty)
                }
                None => Ok(Value::Bool(self.first_style(r)?.align.wrap)),
            },
            "shrinktofit" => match assign {
                Some(v) => {
                    let on = v.to_bool()?;
                    self.restyle(r, move |st| st.align.shrink = on).map(|()| Value::Empty)
                }
                None => Ok(Value::Bool(self.first_style(r)?.align.shrink)),
            },
            "indentlevel" => match assign {
                Some(v) => {
                    let n = u8::try_from(v.to_i32()?.clamp(0, 15)).unwrap_or(0);
                    self.restyle(r, move |st| st.align.indent = n).map(|()| Value::Empty)
                }
                None => Ok(Value::Long(i32::from(self.first_style(r)?.align.indent))),
            },
            "font" => Ok(X::value(XObj::Font(r.clone()))),
            "interior" => Ok(X::value(XObj::Interior(r.clone()))),
            "borders" => match arg(args, 0, "index") {
                None => Ok(X::value(XObj::Borders(r.clone(), None))),
                Some(v) => Ok(X::value(XObj::Borders(r.clone(), Some(v.to_i32()?)))),
            },
            "worksheet" | "parent" => Ok(X::value(XObj::Sheet(r.book, r.sheet.clone()))),
            "select" | "activate" => {
                self.run(r.book, "sheet.activate", json!({"sheet": &*r.sheet}))?;
                self.run(r.book, "selection.set", json!({"range": r.r.a1(), "sheet": &*r.sheet}))?;
                Ok(Value::Empty)
            }
            "delete" => match r.span {
                Span::Rows => self.run_value(r.book, "home.deleteRows", json!({"sheet": &*r.sheet, "rows": r.r.a1()})),
                Span::Cols => self.run_value(r.book, "home.deleteColumns", json!({"sheet": &*r.sheet, "cols": r.r.a1()})),
                Span::Cells => self.run_value(r.book, "home.deleteCells", json!({"sheet": &*r.sheet, "range": r.r.a1(), "shift": "up"})),
            },
            "insert" => match r.span {
                Span::Rows => self.run_value(r.book, "home.insertRows", json!({"sheet": &*r.sheet, "rows": r.r.a1()})),
                Span::Cols => self.run_value(r.book, "home.insertColumns", json!({"sheet": &*r.sheet, "cols": r.r.a1()})),
                Span::Cells => self.run_value(r.book, "home.insertCells", json!({"sheet": &*r.sheet, "range": r.r.a1(), "shift": "down"})),
            },
            _ => Err(VbError::no_member(name)),
        }
    }

    fn simple(&mut self, r: &RangeObj, id: &str) -> Result<Value> {
        self.run_value(r.book, id, Self::range_a1(r))
    }

    fn run_value(&mut self, book: usize, id: &str, params: Json) -> Result<Value> {
        self.run(book, id, params).map(|_| Value::Empty)
    }

    // ---- values

    /// The cells of `r` as a list, row by row (what `For Each` visits).
    pub(super) fn range_cells(&mut self, r: &RangeObj) -> Result<Vec<Value>> {
        let n = r.r.count();
        if n > READ_LIMIT {
            return Err(VbError::new(7, "Out of memory"));
        }
        Ok(r.r.iter().map(|c| self.range_obj(r.book, r.sheet.clone(), RangeRef::cell(c), Span::Cells)).collect())
    }

    fn read_values(&mut self, r: &RangeObj, args: &[CallArg]) -> Result<Value> {
        let _ = args;
        if r.r.is_single()
            && let Some(v) = self.pending_value(r.book, &r.sheet, r.r.start)
        {
            return Ok(v);
        }
        self.flush()?;
        let n = r.r.count();
        if n > READ_LIMIT {
            return Err(VbError::new(7, "Out of memory"));
        }
        let rr = r.r;
        let rows = self.with_sheet(r.book, &r.sheet, |sh, wb| {
            let mut out: Vec<Vec<Value>> = Vec::with_capacity(rr.height() as usize);
            for row in rr.start.row..=rr.end.row {
                let mut line = Vec::with_capacity(rr.width() as usize);
                for col in rr.start.col..=rr.end.col {
                    let at = CellRef::new(row, col);
                    let v = sh.value(at);
                    let date = matches!(v, Cv::Number(_)) && crate::display::parsed(wb.styles.get(sh.style_id(at)).num_fmt.as_str()).is_date();
                    line.push(to_vba(&v, date));
                }
                out.push(line);
            }
            out
        })?;
        if r.r.is_single() {
            return Ok(rows.into_iter().next().and_then(|l| l.into_iter().next()).unwrap_or(Value::Empty));
        }
        let (h, w) = (rows.len() as i32, rr.width() as i32);
        let mut arr = VArray::new(vec![(1, h), (1, w)], Value::Empty)?;
        for (i, line) in rows.into_iter().enumerate() {
            for (j, v) in line.into_iter().enumerate() {
                arr.set(&[i as i32 + 1, j as i32 + 1], v)?;
            }
        }
        Ok(Value::array(arr))
    }

    fn set_values(&mut self, r: &RangeObj, v: &Value) -> Result<()> {
        if let Some(rows) = array_rows(v) {
            let width = rows.iter().map(Vec::len).max().unwrap_or(0) as u64;
            if (rows.len() as u64).saturating_mul(width) > WRITE_LIMIT {
                return Err(VbError::new(7, "Out of memory"));
            }
            for (i, line) in rows.iter().enumerate() {
                for (j, cell) in line.iter().enumerate() {
                    let at =
                        r.r.start
                            .offset(i as i64, j as i64)
                            .filter(|c| c.is_valid())
                            .ok_or_else(|| VbError::app("Application-defined or object-defined error"))?;
                    self.queue_write(r.book, &r.sheet, at, to_json(cell)?)?;
                }
            }
            return Ok(());
        }
        let j = to_json(v)?;
        if r.r.count() > WRITE_LIMIT {
            return Err(VbError::new(7, "Out of memory"));
        }
        for at in r.r.iter() {
            self.queue_write(r.book, &r.sheet, at, j.clone())?;
        }
        Ok(())
    }

    fn set_formula(&mut self, r: &RangeObj, v: &Value) -> Result<()> {
        let text = v.to_text()?;
        if r.r.is_single() {
            return self.queue_write(r.book, &r.sheet, r.r.start, Json::String(text));
        }
        self.flush()?;
        self.run(r.book, "range.fill", json!({"sheet": &*r.sheet, "range": r.r.a1(), "input": text}))?;
        Ok(())
    }

    // ---- font, interior, borders

    pub(super) fn font_member(&mut self, r: &RangeObj, name: &str, assign: Option<Value>) -> Result<Value> {
        match name {
            "bold" => match assign {
                Some(v) => {
                    let on = v.to_bool()?;
                    self.restyle(r, move |st| st.font.bold = on).map(|()| Value::Empty)
                }
                None => Ok(Value::Bool(self.first_style(r)?.font.bold)),
            },
            "italic" => match assign {
                Some(v) => {
                    let on = v.to_bool()?;
                    self.restyle(r, move |st| st.font.italic = on).map(|()| Value::Empty)
                }
                None => Ok(Value::Bool(self.first_style(r)?.font.italic)),
            },
            "strikethrough" => match assign {
                Some(v) => {
                    let on = v.to_bool()?;
                    self.restyle(r, move |st| st.font.strike = on).map(|()| Value::Empty)
                }
                None => Ok(Value::Bool(self.first_style(r)?.font.strike)),
            },
            "underline" => match assign {
                Some(v) => {
                    // xlUnderlineStyleNone is -4142; any other value underlines.
                    let on = v.to_i32()? != -4142;
                    self.restyle(r, move |st| st.font.underline = if on { Underline::Single } else { Underline::None }).map(|()| Value::Empty)
                }
                None => Ok(Value::Long(if self.first_style(r)?.font.underline == Underline::None { -4142 } else { 2 })),
            },
            "size" => match assign {
                Some(v) => {
                    let n = v.to_f64()?;
                    if !(1.0..=409.0).contains(&n) {
                        return Err(VbError::app("Unable to set the Size property of the Font class"));
                    }
                    self.restyle(r, move |st| st.font.size = n as f32).map(|()| Value::Empty)
                }
                None => Ok(Value::Double(f64::from(self.first_style(r)?.font.size))),
            },
            "name" => match assign {
                Some(v) => {
                    let n = v.to_text()?;
                    self.restyle(r, move |st| st.font.name = n.clone()).map(|()| Value::Empty)
                }
                None => Ok(Value::string(self.first_style(r)?.font.name)),
            },
            "color" => match assign {
                Some(v) => {
                    let c = color_of(v.to_i64()?);
                    self.restyle(r, move |st| st.font.color = c).map(|()| Value::Empty)
                }
                None => Ok(vba_color(self.first_style(r)?.font.color)),
            },
            _ => Err(VbError::no_member(name)),
        }
    }

    pub(super) fn interior_member(&mut self, r: &RangeObj, name: &str, assign: Option<Value>) -> Result<Value> {
        match name {
            "color" | "patterncolor" => match assign {
                Some(v) => {
                    let c = color_of(v.to_i64()?);
                    self.restyle(r, move |st| st.fill = Fill::solid(c)).map(|()| Value::Empty)
                }
                None => Ok(vba_color(self.first_style(r)?.fill.fg)),
            },
            "colorindex" | "pattern" => match assign {
                // xlNone (-4142) and xlColorIndexNone clear the fill; any other index keeps it.
                Some(v) => {
                    if v.to_i32()? == -4142 {
                        self.restyle(r, |st| st.fill = Fill::default())?;
                    }
                    Ok(Value::Empty)
                }
                None => {
                    let none = self.first_style(r)?.fill == Fill::default();
                    Ok(Value::Long(if none { -4142 } else { 1 }))
                }
            },
            _ => Err(VbError::no_member(name)),
        }
    }

    /// The border lines of `r` that an index names, as the strips of cells each belongs to.
    fn border_strips(r: &RangeObj, index: Option<i32>) -> Result<Vec<(RangeRef, Side)>> {
        let a = r.r;
        let top = RangeRef::new(a.start, CellRef::new(a.start.row, a.end.col));
        let bottom = RangeRef::new(CellRef::new(a.end.row, a.start.col), a.end);
        let left = RangeRef::new(a.start, CellRef::new(a.end.row, a.start.col));
        let right = RangeRef::new(CellRef::new(a.start.row, a.end.col), a.end);
        let mut out = Vec::new();
        let inside_h = |out: &mut Vec<(RangeRef, Side)>| {
            if a.height() > 1 {
                out.push((RangeRef::new(a.start, CellRef::new(a.end.row - 1, a.end.col)), Side::Bottom));
                out.push((RangeRef::new(CellRef::new(a.start.row + 1, a.start.col), a.end), Side::Top));
            }
        };
        let inside_v = |out: &mut Vec<(RangeRef, Side)>| {
            if a.width() > 1 {
                out.push((RangeRef::new(a.start, CellRef::new(a.end.row, a.end.col - 1)), Side::Right));
                out.push((RangeRef::new(CellRef::new(a.start.row, a.start.col + 1), a.end), Side::Left));
            }
        };
        match index {
            None => {
                out.extend([(top, Side::Top), (bottom, Side::Bottom), (left, Side::Left), (right, Side::Right)]);
                inside_h(&mut out);
                inside_v(&mut out);
            }
            Some(7) => out.push((left, Side::Left)),
            Some(8) => out.push((top, Side::Top)),
            Some(9) => out.push((bottom, Side::Bottom)),
            Some(10) => out.push((right, Side::Right)),
            Some(11) => inside_v(&mut out),
            Some(12) => inside_h(&mut out),
            Some(5) => out.push((a, Side::DiagDown)),
            Some(6) => out.push((a, Side::DiagUp)),
            Some(_) => return Err(VbError::subscript()),
        }
        Ok(out)
    }

    pub(super) fn borders_member(&mut self, r: &RangeObj, index: Option<i32>, name: &str, args: &[CallArg], assign: Option<Value>) -> Result<Value> {
        // `Borders(xlEdgeLeft)` reached through the collection's default member.
        if name.is_empty() || name == "item" {
            let i = arg(args, 0, "index").ok_or_else(VbError::invalid_call)?.to_i32()?;
            return Ok(X::value(XObj::Borders(r.clone(), Some(i))));
        }
        let e = match (name, &assign) {
            ("linestyle", Some(v)) => Edit::Line(v.to_i32()?),
            ("weight", Some(v)) => Edit::Weight(v.to_i32()?),
            ("color", Some(v)) => Edit::Color(color_of(v.to_i64()?)),
            ("colorindex", Some(_)) | ("tintandshade", Some(_)) | ("themecolor", Some(_)) => return Ok(Value::Empty),
            ("linestyle" | "weight" | "color", None) => {
                let strips = Self::border_strips(r, index)?;
                let Some((_, side)) = strips.first().copied() else { return Ok(Value::Empty) };
                let mut style = self.first_style(r)?;
                let line = *line_mut(&mut style.borders, side);
                return Ok(match name {
                    "linestyle" => Value::Long(line_style_value(line.style)),
                    "weight" => Value::Long(weight_value(line.style)),
                    _ => vba_color(line.color),
                });
            }
            _ => return Err(VbError::no_member(name)),
        };
        for (area, side) in Self::border_strips(r, index)? {
            let sub = RangeObj { book: r.book, sheet: r.sheet.clone(), r: area, span: Span::Cells };
            self.restyle(&sub, move |st| {
                let l = line_mut(&mut st.borders, side);
                *l = edited(*l, e);
            })?;
        }
        Ok(Value::Empty)
    }
}

/// `A1`-style address of `r` as `Range.Address` writes it: absolute unless the arguments say not.
fn address(r: &RangeRef, args: &[CallArg]) -> String {
    let row_abs = arg(args, 0, "rowabsolute").and_then(|v| v.to_bool().ok()).unwrap_or(true);
    let col_abs = arg(args, 1, "columnabsolute").and_then(|v| v.to_bool().ok()).unwrap_or(true);
    let cell = |c: CellRef| format!("{}{}{}{}", if col_abs { "$" } else { "" }, col_to_letters(c.col), if row_abs { "$" } else { "" }, c.row + 1);
    if r.is_single() { cell(r.start) } else { format!("{}:{}", cell(r.start), cell(r.end)) }
}

/// `Range.End(direction)`: where Ctrl+Arrow lands from `start`.
fn end_of(sh: &celdas_model::Sheet, start: CellRef, dr: i64, dc: i64) -> CellRef {
    let empty = |c: CellRef| sh.value(c).is_empty();
    let next = |c: CellRef| c.offset(dr, dc).filter(|n| n.is_valid());
    let mut c = start;
    let Some(n) = next(c) else { return c };
    if !empty(c) && !empty(n) {
        c = n;
        while let Some(n) = next(c).filter(|n| !empty(*n)) {
            c = n;
        }
    } else {
        c = n;
        while empty(c) {
            match next(c) {
                Some(n) => c = n,
                None => break,
            }
        }
    }
    c
}

/// The block of cells around `start` that no empty row and column separate (`CurrentRegion`).
fn current_region(sh: &celdas_model::Sheet, start: CellRef) -> RangeRef {
    let mut r = RangeRef::cell(start);
    loop {
        let mut grown = false;
        let ring =
            |r: &RangeRef| (r.start.row.saturating_sub(1), r.end.row.saturating_add(1), r.start.col.saturating_sub(1), r.end.col.saturating_add(1));
        let (t, b, l, rt) = ring(&r);
        let wider = RangeRef::new(CellRef::new(t, l), CellRef::new(b.min(MAX_ROWS - 1), rt.min(MAX_COLS - 1)));
        // A neighbouring row or column that holds anything next to the block joins it.
        if r.start.row > 0 && (wider.start.col..=wider.end.col).any(|c| !sh.value(CellRef::new(t, c)).is_empty()) {
            r.start.row = t;
            grown = true;
        }
        if r.end.row + 1 < MAX_ROWS && (wider.start.col..=wider.end.col).any(|c| !sh.value(CellRef::new(b, c)).is_empty()) {
            r.end.row = b;
            grown = true;
        }
        if r.start.col > 0 && (r.start.row..=r.end.row).any(|row| !sh.value(CellRef::new(row, l)).is_empty()) {
            r.start.col = l;
            grown = true;
        }
        if r.end.col + 1 < MAX_COLS && (r.start.row..=r.end.row).any(|row| !sh.value(CellRef::new(row, rt)).is_empty()) {
            r.end.col = rt;
            grown = true;
        }
        if !grown {
            return r;
        }
    }
}
