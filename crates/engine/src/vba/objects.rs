//! The objects a VBA macro reaches: the application, workbooks, sheets, ranges and drawings. Each
//! one is a small handle (which workbook, which sheet, which cells); the state lives in the
//! workbook, so a handle never goes stale in a way a name lookup can't recover from.

use std::any::Any;
use std::rc::Rc;

use celdas_core::RangeRef;
use celdas_vba::{HostObject, Value};

/// How a range was reached, which decides what `Count`, `AutoFit` and `Delete` mean for it.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum Span {
    /// `Range(...)`, `Cells(...)`.
    Cells,
    /// `.Rows(...)`: its rows.
    Rows,
    /// `.Columns(...)`: its columns.
    Cols,
}

/// A block of cells on a sheet of one of the macro's workbooks.
#[derive(Clone, Debug, PartialEq)]
pub struct RangeObj {
    pub book: usize,
    pub sheet: Rc<str>,
    pub r: RangeRef,
    pub span: Span,
}

/// A shape or chart on a sheet (they share one id space in a workbook).
#[derive(Clone, Debug, PartialEq)]
pub struct ShapeKey {
    pub book: usize,
    pub sheet: Rc<str>,
    pub id: u32,
}

#[derive(Clone, Debug, PartialEq)]
pub enum XObj {
    Application,
    WorksheetFunction,
    Workbooks,
    Shell,
    Workbook(usize),
    Sheets(usize),
    Sheet(usize, Rc<str>),
    Range(RangeObj),
    Font(RangeObj),
    Interior(RangeObj),
    /// The borders of a range: all of them, or the one with this `xlBordersIndex`.
    Borders(RangeObj, Option<i32>),
    Shapes(usize, Rc<str>),
    Shape(ShapeKey),
    TextFrame(ShapeKey),
    TextRange(ShapeKey),
    Chart(ShapeKey),
    SeriesCollection(ShapeKey),
    Series(ShapeKey, usize),
    ChartTitle(ShapeKey),
}

/// The value the interpreter holds for an [`XObj`].
#[derive(Clone, Debug, PartialEq)]
pub struct X(pub XObj);

impl X {
    pub fn value(o: XObj) -> Value {
        Value::object(X(o))
    }
}

impl HostObject for X {
    fn class(&self) -> String {
        match &self.0 {
            XObj::Application => "Application",
            XObj::WorksheetFunction => "WorksheetFunction",
            XObj::Workbooks => "Workbooks",
            XObj::Shell => "IWshShell3",
            XObj::Workbook(_) => "Workbook",
            XObj::Sheets(_) => "Sheets",
            XObj::Sheet(..) => "Worksheet",
            XObj::Range(_) => "Range",
            XObj::Font(_) => "Font",
            XObj::Interior(_) => "Interior",
            XObj::Borders(_, None) => "Borders",
            XObj::Borders(_, Some(_)) => "Border",
            XObj::Shapes(..) => "Shapes",
            XObj::Shape(_) => "Shape",
            XObj::TextFrame(_) => "TextFrame",
            XObj::TextRange(_) => "TextRange",
            XObj::Chart(_) => "Chart",
            XObj::SeriesCollection(_) => "SeriesCollection",
            XObj::Series(..) => "Series",
            XObj::ChartTitle(_) => "ChartTitle",
        }
        .to_string()
    }
    fn as_any(&self) -> &dyn Any {
        self
    }
    fn same(&self, other: &dyn HostObject) -> bool {
        other.as_any().downcast_ref::<X>().is_some_and(|o| o == self)
    }
}
