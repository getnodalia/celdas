//! Drawings: the shapes of a sheet, their text, and the charts a macro builds.

use std::rc::Rc;

use celdas_core::RangeRef;
use celdas_model::{Anchor, Chart, ChartKind, LegendPos, ObjKind, ObjRef, Series, Shape, ShapeKind, Sheet};
use celdas_vba::{CallArg, Value, VbError, arg};

use super::host::{ExcelHost, engine_err, same_name};
use super::objects::{ShapeKey, X, XObj};
use crate::EngineError;

type Result<T> = std::result::Result<T, VbError>;

const PROTECTED: &str = "The cell or chart you're trying to change is on a protected sheet.";

fn other(msg: &str) -> EngineError {
    EngineError::Other(msg.to_string())
}

/// The name VBA shows for a drawing nobody named.
fn default_name(kind: Option<ShapeKind>, id: u32) -> String {
    let label = match kind {
        None => "Chart",
        Some(ShapeKind::TextBox) => "TextBox",
        Some(ShapeKind::Ellipse) => "Oval",
        Some(ShapeKind::RoundedRectangle) => "Rounded Rectangle",
        Some(ShapeKind::Rectangle) => "Rectangle",
        Some(ShapeKind::Line) => "Straight Connector",
        Some(ShapeKind::Arrow) => "Arrow",
        Some(_) => "Shape",
    };
    format!("{label} {id}")
}

/// What a sheet holds, for `Shapes`: shapes first, then charts.
fn drawings(sh: &Sheet) -> Vec<(u32, String, Option<ShapeKind>)> {
    let shapes = sh.shapes.iter().map(|s| (s.id, if s.name.is_empty() { default_name(Some(s.kind), s.id) } else { s.name.clone() }, Some(s.kind)));
    let charts = sh.charts.iter().map(|c| (c.id, if c.name.is_empty() { default_name(None, c.id) } else { c.name.clone() }, None));
    shapes.chain(charts).collect()
}

/// The chart type a VBA `xlChartType` number is.
fn chart_kind(xl: i32) -> (ChartKind, bool) {
    match xl {
        -4169 => (ChartKind::Scatter, false),
        72 | 73 => (ChartKind::ScatterLines, true),
        74 | 75 => (ChartKind::ScatterLines, false),
        4 => (ChartKind::Line, false),
        65 => (ChartKind::LineMarkers, false),
        52 => (ChartKind::ColumnStacked, false),
        53 => (ChartKind::ColumnStacked100, false),
        57 => (ChartKind::BarClustered, false),
        58 => (ChartKind::BarStacked, false),
        59 => (ChartKind::BarStacked100, false),
        5 => (ChartKind::Pie, false),
        -4120 => (ChartKind::Doughnut, false),
        1 => (ChartKind::Area, false),
        76 => (ChartKind::AreaStacked, false),
        -4151 => (ChartKind::Radar, false),
        15 => (ChartKind::Bubble, false),
        _ => (ChartKind::ColumnClustered, false),
    }
}

fn kind_number(k: ChartKind) -> i32 {
    match k {
        ChartKind::Scatter => -4169,
        ChartKind::ScatterLines => 74,
        ChartKind::Line => 4,
        ChartKind::LineMarkers => 65,
        ChartKind::ColumnStacked => 52,
        ChartKind::ColumnStacked100 => 53,
        ChartKind::BarClustered => 57,
        ChartKind::BarStacked => 58,
        ChartKind::BarStacked100 => 59,
        ChartKind::Pie => 5,
        ChartKind::Doughnut => -4120,
        ChartKind::Area => 1,
        ChartKind::AreaStacked => 76,
        ChartKind::Radar => -4151,
        ChartKind::Bubble => 15,
        _ => 51,
    }
}

fn anchor_for(sh: &Sheet, left: f64, top: f64, width: f64, height: f64) -> Anchor {
    let col = sh.col_at(left.max(0.0));
    let row = sh.row_at(top.max(0.0));
    Anchor {
        cell: celdas_core::CellRef::new(row, col),
        dx: (left - sh.col_left(col)).max(0.0) as f32,
        dy: (top - sh.row_top(row)).max(0.0) as f32,
        width: width.clamp(1.0, 10_000.0) as f32,
        height: height.clamp(1.0, 10_000.0) as f32,
        mode: Default::default(),
    }
}

/// A range as a series formula: `'Sheet'!$A$1:$B$2`.
fn abs_ref(sheet: &str, r: &RangeRef) -> String {
    let q = celdas_core::quote_sheet(sheet);
    let col = celdas_core::col_to_letters;
    format!("{q}!${}${}:${}${}", col(r.start.col), r.start.row + 1, col(r.end.col), r.end.row + 1)
}

impl ExcelHost<'_> {
    fn sheet_index(&self, book: usize, sheet: &str) -> Result<usize> {
        self.sheet_names(book)?.iter().position(|n| n == sheet).ok_or_else(VbError::subscript)
    }

    /// Changes sheet `sheet` of `book` as one edit; a protected sheet refuses.
    fn edit_sheet<R>(&mut self, book: usize, sheet: &str, f: impl FnOnce(&mut Sheet, u32) -> std::result::Result<R, EngineError>) -> Result<R> {
        self.flush()?;
        let i = self.sheet_index(book, sheet)?;
        let s = self.session(book)?;
        crate::cmd::edit(s, |cx| {
            let next = cx.wb.next_object_id();
            let sh = cx.sheet_mut(i)?;
            if sh.is_protected() {
                return Err(other(PROTECTED));
            }
            f(sh, next)
        })
        .map_err(engine_err)
    }

    fn find_drawing(&mut self, k: &ShapeKey) -> Result<(String, Option<ShapeKind>)> {
        self.flush()?;
        let id = k.id;
        let found = self.with_sheet(k.book, &k.sheet, |sh, _| drawings(sh).into_iter().find(|d| d.0 == id))?;
        found.map(|d| (d.1, d.2)).ok_or_else(|| VbError::app("The shape no longer exists."))
    }

    pub(super) fn shape_list(&mut self, book: usize, sheet: &Rc<str>) -> Result<Vec<Value>> {
        self.flush()?;
        let all = self.with_sheet(book, sheet, |sh, _| drawings(sh))?;
        Ok(all.into_iter().map(|d| X::value(XObj::Shape(ShapeKey { book, sheet: sheet.clone(), id: d.0 }))).collect())
    }

    pub(super) fn shapes_member(&mut self, book: usize, sheet: &Rc<str>, name: &str, args: &[CallArg]) -> Result<Value> {
        match name {
            "count" => {
                self.flush()?;
                Ok(Value::Long(self.with_sheet(book, sheet, |sh, _| drawings(sh).len())? as i32))
            }
            "" | "item" => {
                self.flush()?;
                let all = self.with_sheet(book, sheet, |sh, _| drawings(sh))?;
                let hit = match arg(args, 0, "index").ok_or_else(VbError::invalid_call)? {
                    Value::Str(s) => all.iter().find(|d| same_name(&d.1, s)),
                    v => {
                        let i = usize::try_from(v.to_i32()?).ok().and_then(|i| i.checked_sub(1)).ok_or_else(VbError::subscript)?;
                        all.get(i)
                    }
                };
                let d = hit.ok_or_else(|| VbError::new(-2147024809, "The item with the specified name wasn't found."))?;
                Ok(X::value(XObj::Shape(ShapeKey { book, sheet: sheet.clone(), id: d.0 })))
            }
            "addchart2" | "addchart" => {
                // AddChart2(Style, Type, Left, Top, Width, Height); AddChart(Type, Left, Top, Width, Height).
                let off = usize::from(name == "addchart2");
                let xl = arg(args, off, "type").map(Value::to_i32).transpose()?.unwrap_or(51);
                let num = |i: usize, n: &str, d: f64| -> Result<f64> { Ok(arg(args, i, n).map(Value::to_f64).transpose()?.unwrap_or(d)) };
                let (left, top) = (num(off + 1, "left", 0.0)?, num(off + 2, "top", 0.0)?);
                let (width, height) = (num(off + 3, "width", 480.0)?, num(off + 4, "height", 288.0)?);
                let (kind, smooth) = chart_kind(xl);
                let id = self.edit_sheet(book, sheet, |sh, id| {
                    let anchor = anchor_for(sh, left, top, width, height);
                    sh.charts.push(Chart {
                        id,
                        kind,
                        anchor,
                        title: None,
                        series: Vec::new(),
                        legend: if matches!(kind, ChartKind::Pie | ChartKind::Doughnut) { LegendPos::Right } else { LegendPos::Bottom },
                        data_labels: false,
                        data_table: None,
                        gridlines: !matches!(kind, ChartKind::Pie | ChartKind::Doughnut),
                        style: 1,
                        hole_size: 75,
                        first_slice_angle: 0,
                        x_title: None,
                        y_title: None,
                        source: None,
                        by_rows: false,
                        pivot: None,
                        name: String::new(),
                    });
                    Ok(id)
                })?;
                if smooth {
                    self.smooth.insert((book, sheet.clone(), id));
                }
                Ok(X::value(XObj::Shape(ShapeKey { book, sheet: sheet.clone(), id })))
            }
            "addtextbox" | "addshape" => {
                // AddTextbox(Orientation, L, T, W, H); AddShape(Type, L, T, W, H).
                let num = |i: usize, n: &str| -> Result<f64> { Ok(arg(args, i, n).map(Value::to_f64).transpose()?.unwrap_or(100.0)) };
                let (left, top, width, height) = (num(1, "left")?, num(2, "top")?, num(3, "width")?, num(4, "height")?);
                let kind = if name == "addtextbox" {
                    ShapeKind::TextBox
                } else {
                    match arg(args, 0, "type").map(Value::to_i32).transpose()?.unwrap_or(1) {
                        9 => ShapeKind::Ellipse,
                        5 => ShapeKind::RoundedRectangle,
                        7 => ShapeKind::Triangle,
                        _ => ShapeKind::Rectangle,
                    }
                };
                let id = self.edit_sheet(book, sheet, |sh, id| {
                    let anchor = anchor_for(sh, left, top, width, height);
                    sh.shapes.push(Shape {
                        id,
                        kind,
                        anchor,
                        fill: celdas_model::Color::Auto,
                        line: celdas_model::Color::Auto,
                        text: String::new(),
                        format: Default::default(),
                        math: Default::default(),
                        name: String::new(),
                        macro_name: String::new(),
                    });
                    Ok(id)
                })?;
                Ok(X::value(XObj::Shape(ShapeKey { book, sheet: sheet.clone(), id })))
            }
            _ => Err(VbError::no_member(name)),
        }
    }

    pub(super) fn shape_member(&mut self, k: &ShapeKey, name: &str, _args: &[CallArg], assign: Option<Value>) -> Result<Value> {
        let (current, kind) = self.find_drawing(k)?;
        let id = k.id;
        match name {
            "name" => match assign {
                Some(v) => {
                    let n = v.to_text()?;
                    self.edit_sheet(k.book, &k.sheet, |sh, _| {
                        if let Some(s) = sh.shapes.iter_mut().find(|s| s.id == id) {
                            s.name = n;
                        } else if let Some(c) = sh.charts.iter_mut().find(|c| c.id == id) {
                            c.name = n;
                        }
                        Ok(())
                    })?;
                    Ok(Value::Empty)
                }
                None => Ok(Value::string(current)),
            },
            "delete" => {
                let kind_ref = if kind.is_some() { ObjKind::Shape } else { ObjKind::Chart };
                self.edit_sheet(k.book, &k.sheet, |sh, _| {
                    sh.remove_object(ObjRef::new(kind_ref, id));
                    Ok(())
                })?;
                Ok(Value::Empty)
            }
            "textframe" | "textframe2" => match kind {
                Some(_) => Ok(X::value(XObj::TextFrame(k.clone()))),
                None => Err(VbError::no_member(name)),
            },
            "chart" => match kind {
                None => Ok(X::value(XObj::Chart(k.clone()))),
                Some(_) => Err(VbError::no_member(name)),
            },
            "onaction" => match assign {
                Some(v) => {
                    let m = v.to_text()?;
                    let m = m.rsplit('!').next().unwrap_or("").to_string();
                    self.edit_sheet(k.book, &k.sheet, |sh, _| {
                        if let Some(s) = sh.shapes.iter_mut().find(|s| s.id == id) {
                            s.macro_name = m;
                        }
                        Ok(())
                    })?;
                    Ok(Value::Empty)
                }
                None => {
                    let m = self.with_sheet(k.book, &k.sheet, |sh, _| sh.shapes.iter().find(|s| s.id == id).map(|s| s.macro_name.clone()))?;
                    Ok(Value::string(m.unwrap_or_default()))
                }
            },
            "type" => Ok(Value::Long(if kind.is_none() {
                3
            } else if kind == Some(ShapeKind::TextBox) {
                17
            } else {
                1
            })),
            "left" | "top" | "width" | "height" => {
                if let Some(v) = assign {
                    let n = v.to_f64()?;
                    self.edit_sheet(k.book, &k.sheet, |sh, _| {
                        let (cl, rt) = {
                            let a = sh.object_ref(id).and_then(|o| sh.object_anchor(o));
                            let (l, t) =
                                a.map_or((0.0, 0.0), |a| (sh.col_left(a.cell.col) + f64::from(a.dx), sh.row_top(a.cell.row) + f64::from(a.dy)));
                            (l, t)
                        };
                        let Some(o) = sh.object_ref(id) else { return Ok(()) };
                        let Some(a) = sh.object_anchor(o) else { return Ok(()) };
                        let (mut l, mut t, mut w, mut h) = (cl, rt, f64::from(a.width), f64::from(a.height));
                        match name {
                            "left" => l = n,
                            "top" => t = n,
                            "width" => w = n,
                            _ => h = n,
                        }
                        let new = anchor_for(sh, l, t, w, h);
                        if let Some(slot) = sh.object_anchor_mut(o) {
                            *slot = Anchor { mode: a.mode, ..new };
                        }
                        Ok(())
                    })?;
                    return Ok(Value::Empty);
                }
                let v = self.with_sheet(k.book, &k.sheet, |sh, _| {
                    let a = sh.object_ref(id).and_then(|o| sh.object_anchor(o)).unwrap_or_default();
                    match name {
                        "left" => sh.col_left(a.cell.col) + f64::from(a.dx),
                        "top" => sh.row_top(a.cell.row) + f64::from(a.dy),
                        "width" => f64::from(a.width),
                        _ => f64::from(a.height),
                    }
                })?;
                Ok(Value::Double(v))
            }
            "visible" => Ok(Value::Bool(true)),
            "select" | "activate" | "bringtofront" | "sendtoback" | "zorder" => Ok(Value::Empty),
            _ => Err(VbError::no_member(name)),
        }
    }

    pub(super) fn text_frame_member(&mut self, k: &ShapeKey, name: &str, _args: &[CallArg], assign: Option<Value>) -> Result<Value> {
        match name {
            "characters" | "textrange" => Ok(X::value(XObj::TextRange(k.clone()))),
            "text" => self.text_range_member(k, "text", &[], assign),
            _ => Err(VbError::no_member(name)),
        }
    }

    pub(super) fn text_range_member(&mut self, k: &ShapeKey, name: &str, _args: &[CallArg], assign: Option<Value>) -> Result<Value> {
        let id = k.id;
        match name {
            "text" | "caption" => match assign {
                Some(v) => {
                    let t = v.to_text()?;
                    self.edit_sheet(k.book, &k.sheet, |sh, _| {
                        let s = sh.shapes.iter_mut().find(|s| s.id == id).ok_or_else(|| other("The shape no longer exists."))?;
                        s.text = t;
                        Ok(())
                    })?;
                    Ok(Value::Empty)
                }
                None => {
                    self.flush()?;
                    let t = self.with_sheet(k.book, &k.sheet, |sh, _| sh.shapes.iter().find(|s| s.id == id).map(|s| s.text.clone()))?;
                    Ok(Value::string(t.unwrap_or_default()))
                }
            },
            "characters" | "textrange" => Ok(X::value(XObj::TextRange(k.clone()))),
            "length" => {
                let t = self.text_range_member(k, "text", &[], None)?.to_text()?;
                Ok(Value::Long(t.chars().count() as i32))
            }
            _ => Err(VbError::no_member(name)),
        }
    }

    // ---- charts

    fn edit_chart<R>(&mut self, k: &ShapeKey, f: impl FnOnce(&mut Chart) -> std::result::Result<R, EngineError>) -> Result<R> {
        let id = k.id;
        self.edit_sheet(k.book, &k.sheet, |sh, _| {
            let c = sh.charts.iter_mut().find(|c| c.id == id).ok_or_else(|| other("The chart no longer exists."))?;
            f(c)
        })
    }

    fn read_chart<R>(&mut self, k: &ShapeKey, f: impl FnOnce(&Chart) -> R) -> Result<R> {
        self.flush()?;
        let id = k.id;
        self.with_sheet(k.book, &k.sheet, |sh, _| sh.charts.iter().find(|c| c.id == id).map(f))?
            .ok_or_else(|| VbError::app("The chart no longer exists."))
    }

    pub(super) fn chart_member(&mut self, k: &ShapeKey, name: &str, args: &[CallArg], assign: Option<Value>) -> Result<Value> {
        match name {
            "hastitle" => match assign {
                Some(v) => {
                    let on = v.to_bool()?;
                    self.edit_chart(k, |c| {
                        if !on {
                            c.title = None;
                        } else if c.title.is_none() {
                            c.title = Some("Chart Title".into());
                        }
                        Ok(())
                    })?;
                    Ok(Value::Empty)
                }
                None => Ok(Value::Bool(self.read_chart(k, |c| c.title.is_some())?)),
            },
            "charttitle" => Ok(X::value(XObj::ChartTitle(k.clone()))),
            "haslegend" => match assign {
                Some(v) => {
                    let on = v.to_bool()?;
                    self.edit_chart(k, |c| {
                        c.legend = if !on {
                            LegendPos::None
                        } else if c.legend == LegendPos::None {
                            LegendPos::Bottom
                        } else {
                            c.legend
                        };
                        Ok(())
                    })?;
                    Ok(Value::Empty)
                }
                None => Ok(Value::Bool(self.read_chart(k, |c| c.legend != LegendPos::None)?)),
            },
            "charttype" => match assign {
                Some(v) => {
                    let (kind, smooth) = chart_kind(v.to_i32()?);
                    self.edit_chart(k, |c| {
                        c.kind = kind;
                        for s in &mut c.series {
                            s.smooth = smooth;
                        }
                        Ok(())
                    })?;
                    let key = (k.book, k.sheet.clone(), k.id);
                    if smooth {
                        self.smooth.insert(key);
                    } else {
                        self.smooth.remove(&key);
                    }
                    Ok(Value::Empty)
                }
                None => Ok(Value::Long(kind_number(self.read_chart(k, |c| c.kind)?))),
            },
            "seriescollection" => match arg(args, 0, "index") {
                None => Ok(X::value(XObj::SeriesCollection(k.clone()))),
                Some(_) => self.series_collection_member(k, "item", args),
            },
            "parent" | "shape" => Ok(X::value(XObj::Shape(k.clone()))),
            "delete" => {
                let id = k.id;
                self.edit_sheet(k.book, &k.sheet, |sh, _| {
                    sh.remove_object(ObjRef::new(ObjKind::Chart, id));
                    Ok(())
                })?;
                Ok(Value::Empty)
            }
            "refresh" | "activate" | "select" => Ok(Value::Empty),
            _ => Err(VbError::no_member(name)),
        }
    }

    pub(super) fn series_collection_member(&mut self, k: &ShapeKey, name: &str, args: &[CallArg]) -> Result<Value> {
        match name {
            "count" => Ok(Value::Long(self.read_chart(k, |c| c.series.len())? as i32)),
            "" | "item" => {
                let n = self.read_chart(k, |c| c.series.len())?;
                let i = usize::try_from(arg(args, 0, "index").ok_or_else(VbError::invalid_call)?.to_i32()?)
                    .ok()
                    .and_then(|i| i.checked_sub(1))
                    .filter(|i| *i < n)
                    .ok_or_else(VbError::subscript)?;
                Ok(X::value(XObj::Series(k.clone(), i)))
            }
            "newseries" => {
                let smooth = self.smooth.contains(&(k.book, k.sheet.clone(), k.id));
                let i = self.edit_chart(k, |c| {
                    c.series.push(Series {
                        name: None,
                        categories: None,
                        values: String::new(),
                        bubble_sizes: None,
                        color: None,
                        secondary: false,
                        kind: None,
                        quartiles: Default::default(),
                        smooth,
                        trendlines: Vec::new(),
                    });
                    Ok(c.series.len() - 1)
                })?;
                Ok(X::value(XObj::Series(k.clone(), i)))
            }
            _ => Err(VbError::no_member(name)),
        }
    }

    /// The series of a chart as `For Each` visits them.
    pub(super) fn series_list(&mut self, k: &ShapeKey) -> Result<Vec<Value>> {
        let n = self.read_chart(k, |c| c.series.len())?;
        Ok((0..n).map(|i| X::value(XObj::Series(k.clone(), i))).collect())
    }

    /// What a range or text assigned to a series property becomes in the chart.
    fn series_source(v: &Value) -> Result<String> {
        match v {
            Value::Object(Some(o)) => match o.downcast::<X>().map(|x| &x.0) {
                Some(XObj::Range(r)) => Ok(abs_ref(&r.sheet, &r.r)),
                _ => Err(VbError::type_mismatch()),
            },
            Value::Str(s) => Ok(s.to_string()),
            _ => Err(VbError::type_mismatch()),
        }
    }

    pub(super) fn series_member(&mut self, k: &ShapeKey, i: usize, name: &str, assign: Option<Value>) -> Result<Value> {
        match name {
            "name" => match assign {
                Some(v) => {
                    let n = match &v {
                        Value::Object(_) => Self::series_source(&v)?,
                        other => other.to_text()?,
                    };
                    self.edit_chart(k, |c| {
                        c.series.get_mut(i).ok_or_else(|| other_err("series"))?.name = Some(n);
                        Ok(())
                    })?;
                    Ok(Value::Empty)
                }
                None => Ok(Value::string(self.read_chart(k, |c| c.series.get(i).and_then(|s| s.name.clone()))?.unwrap_or_default())),
            },
            "xvalues" | "values" => match assign {
                Some(v) => {
                    let src = Self::series_source(&v)?;
                    let x = name == "xvalues";
                    self.edit_chart(k, |c| {
                        let s = c.series.get_mut(i).ok_or_else(|| other_err("series"))?;
                        if x {
                            s.categories = Some(src);
                        } else {
                            s.values = src;
                        }
                        Ok(())
                    })?;
                    Ok(Value::Empty)
                }
                None => Err(VbError::app(format!("Reading the {name} of a series isn't supported here."))),
            },
            "delete" => {
                self.edit_chart(k, |c| {
                    if i < c.series.len() {
                        c.series.remove(i);
                    }
                    Ok(())
                })?;
                Ok(Value::Empty)
            }
            _ => Err(VbError::no_member(name)),
        }
    }

    pub(super) fn chart_title_member(&mut self, k: &ShapeKey, name: &str, assign: Option<Value>) -> Result<Value> {
        match name {
            "text" | "caption" => match assign {
                Some(v) => {
                    let t = v.to_text()?;
                    self.edit_chart(k, |c| {
                        c.title = Some(t);
                        Ok(())
                    })?;
                    Ok(Value::Empty)
                }
                None => Ok(Value::string(self.read_chart(k, |c| c.title.clone())?.unwrap_or_default())),
            },
            _ => Err(VbError::no_member(name)),
        }
    }
}

fn other_err(what: &str) -> EngineError {
    EngineError::Other(format!("The {what} no longer exists."))
}
