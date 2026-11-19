//! Sample workbooks built in code (original content), for the start screen, demos, README
//! screenshots and tests. Each one is written in English or in Spanish: a workbook started from a
//! Spanish interface arrives in Spanish.

use std::sync::Arc;

use celdas_model::{Sheet, Workbook};
use serde_json::json;

use crate::{DocState, Session};

/// The language a sample is written in.
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
pub enum Lang {
    #[default]
    En,
    Es,
}

impl Lang {
    /// `"es"`, `"es-ES"` and `"es_MX.UTF-8"` are Spanish; anything else is English.
    pub fn parse(code: &str) -> Lang {
        let code = code.trim().to_ascii_lowercase();
        if code == "es" || code.starts_with("es-") || code.starts_with("es_") { Lang::Es } else { Lang::En }
    }

    /// `en` or `es`, by language.
    pub fn pick(self, en: &'static str, es: &'static str) -> &'static str {
        match self {
            Lang::En => en,
            Lang::Es => es,
        }
    }
}

/// The samples in gallery order: id and English title. `title_in` gives the Spanish one.
pub const SAMPLES: &[(&str, &str)] = &[
    ("budget", "Household Budget"),
    ("sales", "Quarterly Sales Dashboard"),
    ("grades", "Gradebook"),
    ("invoice", "Invoice"),
    ("tracker", "Project Tracker"),
    ("loan", "Loan Calculator"),
    ("inventory", "Stock Inventory"),
    ("timesheet", "Weekly Timesheet"),
    ("savings", "Savings Goal"),
    ("expenses", "Expense Report"),
    ("planner", "Weekly Planner"),
];

const SPANISH_TITLES: &[(&str, &str)] = &[
    ("budget", "Presupuesto doméstico"),
    ("sales", "Panel de ventas trimestrales"),
    ("grades", "Libro de calificaciones"),
    ("invoice", "Factura"),
    ("tracker", "Seguimiento de proyecto"),
    ("loan", "Calculadora de préstamo"),
    ("inventory", "Inventario de existencias"),
    ("timesheet", "Hoja de horas semanal"),
    ("savings", "Objetivo de ahorro"),
    ("expenses", "Informe de gastos"),
    ("planner", "Planificador semanal"),
];

pub fn title(name: &str) -> String {
    title_in(name, Lang::En)
}

/// The sample's title in `lang`.
pub fn title_in(name: &str, lang: Lang) -> String {
    let list = if lang == Lang::Es { SPANISH_TITLES } else { SAMPLES };
    list.iter().find(|(n, _)| *n == name).map(|(_, t)| t.to_string()).unwrap_or_else(|| lang.pick("Sample", "Ejemplo").into())
}

/// Builds a sample by replaying commands (so samples also exercise the command layer).
pub fn build(name: &str) -> Option<Workbook> {
    build_in(name, Lang::En)
}

/// [`build`] in `lang`.
pub fn build_in(name: &str, lang: Lang) -> Option<Workbook> {
    let mut s = Session::new();
    let mut wb = Workbook::new();
    if let Some(sh) = wb.sheet_mut(0) {
        sh.name = match name {
            "budget" => lang.pick("Budget", "Presupuesto"),
            "sales" => lang.pick("Sales", "Ventas"),
            "grades" => lang.pick("Grades", "Notas"),
            "invoice" => lang.pick("Invoice", "Factura"),
            "tracker" => lang.pick("Tracker", "Seguimiento"),
            "loan" => lang.pick("Loan", "Préstamo"),
            "inventory" => lang.pick("Inventory", "Inventario"),
            "timesheet" => lang.pick("Timesheet", "Horas"),
            "savings" => lang.pick("Savings", "Ahorro"),
            "expenses" => lang.pick("Expenses", "Gastos"),
            "planner" => lang.pick("Planner", "Semana"),
            _ => return None,
        }
        .into();
    }
    s.add_document(DocState::new(wb, None, title_in(name, lang), crate::Trust::Local));
    let ok = match name {
        "budget" => budget(&mut s, lang),
        "sales" => sales(&mut s, lang),
        "grades" => grades(&mut s, lang),
        "invoice" => invoice(&mut s, lang),
        "tracker" => tracker(&mut s, lang),
        "loan" => loan(&mut s, lang),
        "inventory" => inventory(&mut s, lang),
        "timesheet" => timesheet(&mut s, lang),
        "savings" => savings(&mut s, lang),
        "expenses" => expenses(&mut s, lang),
        "planner" => planner(&mut s, lang),
        _ => false,
    };
    if !ok {
        return None;
    }
    let mut wb = (*s.doc().ok()?.wb).clone();
    wb.set_active_sheet(0);
    if let Some(sh) = wb.sheet_mut(0) {
        sh.view_active = celdas_core::CellRef::default();
    }
    Some(wb)
}

fn run(s: &mut Session, id: &str, p: serde_json::Value) -> bool {
    match s.execute(id, p) {
        Ok(_) => true,
        Err(e) => {
            log::warn!("sample: {id} failed: {e}");
            false
        }
    }
}

fn sales(s: &mut Session, lang: Lang) -> bool {
    let (region, total, growth, trend) =
        (lang.pick("Region", "Región"), lang.pick("Total", "Total"), lang.pick("Growth", "Crecimiento"), lang.pick("Trend", "Tendencia"));
    let table = lang.pick("Sales", "Ventas");
    let q = [lang.pick("Q1", "T1"), lang.pick("Q2", "T2"), lang.pick("Q3", "T3"), lang.pick("Q4", "T4")];
    let (north, south, east, west, central, online) = match lang {
        Lang::En => ("North", "South", "East", "West", "Central", "Online"),
        Lang::Es => ("Norte", "Sur", "Este", "Oeste", "Centro", "En línea"),
    };
    let rows = json!([
        [region, q[0], q[1], q[2], q[3], total, growth, trend],
        [north, 48200, 51340, 56010, 61870],
        [south, 39150, 37980, 42400, 45120],
        [east, 55300, 58910, 57230, 63480],
        [west, 61020, 64400, 70150, 76390],
        [central, 28740, 31260, 33980, 36510],
        [online, 72410, 81630, 95220, 112840],
    ]);
    let mut ok = run(s, "range.setValues", json!({"range": "B4", "values": rows}));
    ok &= run(s, "range.setValues", json!({"range": "B2", "values": [[lang.pick("Quarterly Sales Dashboard", "Panel de ventas trimestrales")]]}));
    ok &= run(s, "home.cellStyle", json!({"range": "B2", "name": "Title"}));
    ok &= run(
        s,
        "range.setValues",
        json!({"range": "B3", "values": [[lang.pick("Fiscal year 2026 · all figures in USD", "Ejercicio 2026 · todas las cifras en USD")]]}),
    );
    ok &= run(s, "home.cellStyle", json!({"range": "B3", "name": "Explanatory Text"}));
    for r in 5..=10 {
        ok &= run(s, "cell.set", json!({"cell": format!("G{r}"), "input": format!("=SUM(C{r}:F{r})")}));
        ok &= run(s, "cell.set", json!({"cell": format!("H{r}"), "input": format!("=F{r}/C{r}-1")}));
        ok &= run(s, "insert.sparkline", json!({"range": format!("C{r}:F{r}"), "location": format!("I{r}"), "type": "line", "markers": true}));
    }
    ok &= run(s, "insert.table", json!({"range": "B4:I10", "style": "TableStyleMedium2", "name": table}));
    ok &= run(s, "table.totalRow", json!({"table": table, "on": true}));
    for col in [q[0], q[1], q[2], q[3], total] {
        ok &= run(s, "table.totalFunction", json!({"table": table, "column": col, "function": "sum"}));
    }
    ok &= run(s, "table.totalFunction", json!({"table": table, "column": growth, "function": "average"}));
    ok &= run(s, "table.totalFunction", json!({"table": table, "column": trend, "function": "none"}));
    ok &= run(s, "home.numberFormat", json!({"range": "C5:G11", "code": "\"$\"#,##0"}));
    ok &= run(s, "home.numberFormat", json!({"range": "H5:H11", "code": "0.0%"}));
    ok &= run(s, "home.conditionalFormat", json!({"range": "G5:G10", "rule": {"type": "dataBar", "color": "#5B9BD5"}}));
    ok &= run(s, "home.conditionalFormat", json!({"range": "H5:H10", "rule": {"type": "iconSet", "set": "3Arrows"}}));
    ok &= run(s, "home.columnWidth", json!({"cols": "A:A", "width": 20}));
    ok &= run(s, "home.columnWidth", json!({"cols": "B:B", "width": 132}));
    ok &= run(s, "home.columnWidth", json!({"cols": "C:G", "width": 84}));
    ok &= run(s, "home.columnWidth", json!({"cols": "H:I", "width": 76}));
    ok &= run(s, "selection.set", json!({"range": "B4:F10"}));
    ok &= run(
        s,
        "insert.chart",
        json!({"type": "column", "title": lang.pick("Revenue by Region", "Ingresos por región"), "at": "K3", "width": 460, "height": 270}),
    );
    ok &= run(s, "selection.set", json!({"range": "B4:B10"}));
    // Pie of totals.
    ok &= run(
        s,
        "insert.chart",
        json!({"range": "B4:B10", "type": "pie", "title": lang.pick("Share of Annual Total", "Peso en el total anual"), "at": "K19", "width": 460, "height": 260}),
    );
    if let Ok(d) = s.doc_mut() {
        let wb = Arc::make_mut(&mut d.wb);
        let sheet = lang.pick("Sales", "Ventas");
        if let Some(sh) = wb.sheet_mut(0)
            && let Some(c) = sh.charts.last_mut()
        {
            c.series = vec![celdas_model::Series {
                name: Some(format!("{sheet}!$G$4")),
                categories: Some(format!("{sheet}!$B$5:$B$10")),
                values: format!("{sheet}!$G$5:$G$10"),
                bubble_sizes: None,
                color: None,
                secondary: false,
                kind: None,
                quartiles: Default::default(),
                smooth: false,
                trendlines: Vec::new(),
            }];
        }
    }
    let labels = match lang {
        Lang::En => [["Best region"], ["Average quarter"], ["Online share"], ["Year-over-year"]],
        Lang::Es => [["Mejor región"], ["Trimestre medio"], ["Peso en línea"], ["Interanual"]],
    };
    ok &= run(s, "range.setValues", json!({"range": "B14", "values": labels}));
    ok &= run(s, "cell.set", json!({"cell": "C14", "input": format!("=INDEX({table}[{region}],MATCH(MAX({table}[{total}]),{table}[{total}],0))")}));
    ok &= run(s, "cell.set", json!({"cell": "C15", "input": "=AVERAGE(C5:F10)"}));
    ok &= run(s, "cell.set", json!({"cell": "C16", "input": "=G10/SUM(G5:G10)"}));
    ok &= run(s, "cell.set", json!({"cell": "C17", "input": "=SUM(F5:F10)/SUM(C5:C10)-1"}));
    ok &= run(s, "home.numberFormat", json!({"range": "C15", "code": "\"$\"#,##0"}));
    ok &= run(s, "home.numberFormat", json!({"range": "C16:C17", "code": "0.0%"}));
    ok &= run(s, "home.bold", json!({"range": "B14:B17", "on": true}));
    ok &= run(s, "home.borders", json!({"range": "B14:C17", "preset": "outside"}));
    ok &= run(s, "home.fillColor", json!({"range": "B14:C17", "color": "#F2F7FC"}));
    ok &= run(
        s,
        "review.newNote",
        json!({"cell": "H4", "text": lang.pick("Q4 vs Q1 growth for each region.", "Crecimiento del T4 frente al T1 en cada región.")}),
    );
    ok &= run(s, "selection.set", json!({"cell": "B2"}));
    ok
}

fn budget(s: &mut Session, lang: Lang) -> bool {
    let rows = match lang {
        Lang::En => json!([
            ["Category", "Budget", "Actual", "Difference", "% Used"],
            ["Rent", 1850, 1850],
            ["Groceries", 620, 684.35],
            ["Utilities", 240, 212.8],
            ["Transport", 180, 203.5],
            ["Internet & phone", 95, 95],
            ["Insurance", 210, 210],
            ["Dining out", 200, 265.9],
            ["Entertainment", 120, 88.25],
            ["Savings", 600, 600],
            ["Art supplies", 150, 172.4],
        ]),
        Lang::Es => json!([
            ["Categoría", "Presupuesto", "Real", "Diferencia", "% usado"],
            ["Alquiler", 1850, 1850],
            ["Compra", 620, 684.35],
            ["Suministros", 240, 212.8],
            ["Transporte", 180, 203.5],
            ["Internet y móvil", 95, 95],
            ["Seguros", 210, 210],
            ["Comer fuera", 200, 265.9],
            ["Ocio", 120, 88.25],
            ["Ahorro", 600, 600],
            ["Material de arte", 150, 172.4],
        ]),
    };
    let mut ok = run(
        s,
        "range.setValues",
        json!({"range": "B2", "values": [[lang.pick("Monthly Budget — October 2026", "Presupuesto mensual — octubre de 2026")]]}),
    );
    ok &= run(s, "home.cellStyle", json!({"range": "B2", "name": "Title"}));
    ok &= run(s, "range.setValues", json!({"range": "B4", "values": rows}));
    for r in 5..=14 {
        ok &= run(s, "cell.set", json!({"cell": format!("E{r}"), "input": format!("=C{r}-D{r}")}));
        ok &= run(s, "cell.set", json!({"cell": format!("F{r}"), "input": format!("=D{r}/C{r}")}));
    }
    ok &= run(s, "range.setValues", json!({"range": "B15", "values": [[lang.pick("Total", "Total")]]}));
    for c in ["C", "D", "E"] {
        ok &= run(s, "cell.set", json!({"cell": format!("{c}15"), "input": format!("=SUM({c}5:{c}14)")}));
    }
    ok &= run(s, "cell.set", json!({"cell": "F15", "input": "=D15/C15"}));
    ok &= run(s, "home.cellStyle", json!({"range": "B4:F4", "name": "Heading 3"}));
    ok &= run(s, "home.cellStyle", json!({"range": "B15:F15", "name": "Total"}));
    ok &= run(s, "home.numberFormat", json!({"range": "C5:E15", "format": "Accounting"}));
    ok &= run(s, "home.numberFormat", json!({"range": "F5:F15", "code": "0%"}));
    ok &= run(
        s,
        "home.conditionalFormat",
        json!({"range": "E5:E14", "rule": {"type": "cellIs", "operator": "less", "value": 0, "preset": "lightRedFillDarkRedText"}}),
    );
    ok &= run(s, "home.conditionalFormat", json!({"range": "F5:F14", "rule": {"type": "colorScale", "colors": ["#63BE7B", "#FFEB84", "#F8696B"]}}));
    ok &= run(s, "home.columnWidth", json!({"cols": "B:B", "width": 130}));
    ok &= run(s, "home.columnWidth", json!({"cols": "C:E", "width": 90}));
    ok &= run(s, "selection.set", json!({"range": "B4:D14"}));
    ok &= run(
        s,
        "insert.chart",
        json!({"type": "bar", "title": lang.pick("Budget vs Actual", "Presupuesto frente a real"), "at": "H3", "width": 440, "height": 300}),
    );
    ok &= run(s, "view.freezePanes", json!({"cell": "A5"}));
    ok &= run(s, "selection.set", json!({"cell": "B2"}));
    ok
}

fn grades(s: &mut Session, lang: Lang) -> bool {
    let rows = match lang {
        Lang::En => json!([
            ["Student", "Essay", "Midterm", "Project", "Final", "Average", "Grade"],
            ["Amara Okafor", 92, 88, 95, 91],
            ["Bruno Silva", 78, 82, 74, 80],
            ["Chen Wei", 85, 91, 89, 94],
            ["Dana Kowalski", 67, 72, 80, 70],
            ["Elif Yılmaz", 98, 95, 97, 99],
            ["Farid Haddad", 74, 69, 71, 77],
            ["Grace Mensah", 88, 84, 90, 86],
            ["Hiro Tanaka", 59, 64, 70, 62],
        ]),
        Lang::Es => json!([
            ["Alumno", "Redacción", "Parcial", "Proyecto", "Final", "Media", "Nota"],
            ["Amara Okafor", 92, 88, 95, 91],
            ["Bruno Silva", 78, 82, 74, 80],
            ["Chen Wei", 85, 91, 89, 94],
            ["Dana Kowalski", 67, 72, 80, 70],
            ["Elif Yılmaz", 98, 95, 97, 99],
            ["Farid Haddad", 74, 69, 71, 77],
            ["Grace Mensah", 88, 84, 90, 86],
            ["Hiro Tanaka", 59, 64, 70, 62],
        ]),
    };
    let mut ok = run(s, "range.setValues", json!({"range": "A1", "values": rows}));
    for r in 2..=9 {
        ok &= run(s, "cell.set", json!({"cell": format!("F{r}"), "input": format!("=AVERAGE(B{r}:E{r})")}));
        ok &= run(
            s,
            "cell.set",
            json!({"cell": format!("G{r}"), "input": format!("=IFS(F{r}>=90,\"A\",F{r}>=80,\"B\",F{r}>=70,\"C\",F{r}>=60,\"D\",TRUE,\"F\")")}),
        );
    }
    ok &= run(s, "home.numberFormat", json!({"range": "F2:F9", "code": "0.0"}));
    ok &= run(s, "insert.table", json!({"range": "A1:G9", "style": "TableStyleLight9", "name": lang.pick("Grades", "Notas")}));
    ok &= run(s, "home.columnWidth", json!({"cols": "A:A", "width": 120}));
    ok &= run(
        s,
        "home.conditionalFormat",
        json!({"range": "G2:G9", "rule": {"type": "cellIs", "operator": "equal", "value": "A", "preset": "greenFillDarkGreenText"}}),
    );
    ok &= run(s, "home.conditionalFormat", json!({"range": "F2:F9", "rule": {"type": "dataBar", "color": "#63BE7B"}}));
    ok
}

/// A title in B2 and an optional line under it, as every sample below starts.
fn heading(s: &mut Session, title: &str, subtitle: &str) -> bool {
    let mut ok = run(s, "range.setValues", json!({"range": "B2", "values": [[title]]}));
    ok &= run(s, "home.cellStyle", json!({"range": "B2", "name": "Title"}));
    if !subtitle.is_empty() {
        ok &= run(s, "range.setValues", json!({"range": "B3", "values": [[subtitle]]}));
        ok &= run(s, "home.cellStyle", json!({"range": "B3", "name": "Explanatory Text"}));
    }
    ok
}

/// An invoice: who and what, three line items, VAT at a rate kept in a cell, and the total due.
fn invoice(s: &mut Session, lang: Lang) -> bool {
    let mut ok = heading(s, lang.pick("Invoice", "Factura"), lang.pick("No. 2026-014 · due in 30 days", "N.º 2026-014 · vence a 30 días"));
    let details = match lang {
        Lang::En => [["From", "Alex Rivera · Illustration"], ["To", "Harbor Books"], ["Date", "2026-10-31"]],
        Lang::Es => [["De", "Alejandra Rivera · Ilustración"], ["Para", "Editorial Puerto"], ["Fecha", "2026-10-31"]],
    };
    ok &= run(s, "range.setValues", json!({"range": "B5", "values": details}));
    ok &= run(s, "home.bold", json!({"range": "B5:B7", "on": true}));
    let rows = match lang {
        Lang::En => json!([
            ["Description", "Quantity", "Unit price", "Amount"],
            ["Cover illustration", 1, 400],
            ["Chapter openers", 6, 80],
            ["Revisions (hours)", 2, 45],
        ]),
        Lang::Es => json!([
            ["Concepto", "Cantidad", "Precio unitario", "Importe"],
            ["Ilustración de portada", 1, 400],
            ["Capitulares", 6, 80],
            ["Revisiones (horas)", 2, 45],
        ]),
    };
    ok &= run(s, "range.setValues", json!({"range": "B9", "values": rows}));
    for r in 10..=12 {
        ok &= run(s, "cell.set", json!({"cell": format!("E{r}"), "input": format!("=C{r}*D{r}")}));
    }
    ok &= run(
        s,
        "range.setValues",
        json!({"range": "D14", "values": [[lang.pick("Subtotal", "Base imponible")], [lang.pick("VAT rate", "Tipo de IVA")], [lang.pick("VAT", "IVA")], [lang.pick("Total due", "Total a pagar")]]}),
    );
    ok &= run(s, "cell.set", json!({"cell": "E14", "input": "=SUM(E10:E12)"}));
    ok &= run(s, "cell.set", json!({"cell": "E15", "input": "0.21"}));
    ok &= run(s, "cell.set", json!({"cell": "E16", "input": "=E14*E15"}));
    ok &= run(s, "cell.set", json!({"cell": "E17", "input": "=E14+E16"}));
    ok &= run(s, "home.cellStyle", json!({"range": "B9:E9", "name": "Heading 3"}));
    ok &= run(s, "home.cellStyle", json!({"range": "D17:E17", "name": "Total"}));
    ok &= run(s, "home.numberFormat", json!({"range": "D10:E14", "code": "#,##0.00"}));
    ok &= run(s, "home.numberFormat", json!({"range": "E15", "code": "0%"}));
    ok &= run(s, "home.numberFormat", json!({"range": "E16:E17", "code": "#,##0.00"}));
    ok &= run(s, "home.columnWidth", json!({"cols": "A:A", "width": 20}));
    ok &= run(s, "home.columnWidth", json!({"cols": "B:B", "width": 200}));
    ok &= run(s, "home.columnWidth", json!({"cols": "C:E", "width": 110}));
    ok &= run(s, "selection.set", json!({"cell": "B2"}));
    ok
}

/// A task list with an owner, a status chosen from a list, a week and a progress bar, and a line
/// that counts what is done.
fn tracker(s: &mut Session, lang: Lang) -> bool {
    let (todo, doing, done) = (lang.pick("To do", "Pendiente"), lang.pick("In progress", "En curso"), lang.pick("Done", "Hecho"));
    let mut ok = heading(
        s,
        lang.pick("Project tracker", "Seguimiento de proyecto"),
        lang.pick("Change a status from the list and the summary follows", "Cambia un estado de la lista y el resumen lo sigue"),
    );
    let rows = match lang {
        Lang::En => json!([
            ["Task", "Owner", "Status", "Week", "Progress"],
            ["Agree the goal", "Mara", done, 1, 1],
            ["Draw the first sketches", "Theo", done, 1, 1],
            ["Build the prototype", "Inés", doing, 2, 0.6],
            ["Try it with five people", "Samir", doing, 3, 0.25],
            ["Write down what we learned", "Mara", todo, 4, 0],
            ["Plan the next round", "Theo", todo, 4, 0],
        ]),
        Lang::Es => json!([
            ["Tarea", "Responsable", "Estado", "Semana", "Avance"],
            ["Acordar el objetivo", "Mara", done, 1, 1],
            ["Dibujar los primeros bocetos", "Theo", done, 1, 1],
            ["Construir el prototipo", "Inés", doing, 2, 0.6],
            ["Probarlo con cinco personas", "Samir", doing, 3, 0.25],
            ["Escribir lo aprendido", "Mara", todo, 4, 0],
            ["Planificar la siguiente ronda", "Theo", todo, 4, 0],
        ]),
    };
    ok &= run(s, "range.setValues", json!({"range": "B5", "values": rows}));
    ok &= run(s, "home.cellStyle", json!({"range": "B5:F5", "name": "Heading 3"}));
    ok &= run(s, "data.validation", json!({"range": "D6:D11", "type": "list", "formula1": format!("\"{todo},{doing},{done}\""), "dropdown": true}));
    ok &= run(s, "home.numberFormat", json!({"range": "F6:F11", "code": "0%"}));
    ok &= run(s, "home.conditionalFormat", json!({"range": "F6:F11", "rule": {"type": "dataBar", "color": "#63BE7B"}}));
    ok &= run(
        s,
        "home.conditionalFormat",
        json!({"range": "D6:D11", "rule": {"type": "cellIs", "operator": "equal", "value": format!("\"{done}\""), "preset": "greenFillDarkGreenText"}}),
    );
    ok &= run(
        s,
        "range.setValues",
        json!({"range": "H5", "values": [[lang.pick("Summary", "Resumen")], [lang.pick("Tasks", "Tareas")], [done], [doing], [todo], [lang.pick("Overall progress", "Avance total")]]}),
    );
    ok &= run(s, "home.cellStyle", json!({"range": "H5:I5", "name": "Heading 3"}));
    ok &= run(s, "cell.set", json!({"cell": "I6", "input": "=COUNTA(B6:B11)"}));
    ok &= run(s, "cell.set", json!({"cell": "I7", "input": "=COUNTIF(D6:D11,H7)"}));
    ok &= run(s, "cell.set", json!({"cell": "I8", "input": "=COUNTIF(D6:D11,H8)"}));
    ok &= run(s, "cell.set", json!({"cell": "I9", "input": "=COUNTIF(D6:D11,H9)"}));
    ok &= run(s, "cell.set", json!({"cell": "I10", "input": "=AVERAGE(F6:F11)"}));
    ok &= run(s, "home.numberFormat", json!({"range": "I10", "code": "0%"}));
    ok &= run(s, "home.columnWidth", json!({"cols": "A:A", "width": 20}));
    ok &= run(s, "home.columnWidth", json!({"cols": "B:B", "width": 210}));
    ok &= run(s, "home.columnWidth", json!({"cols": "C:F", "width": 96}));
    ok &= run(s, "home.columnWidth", json!({"cols": "H:H", "width": 120}));
    ok &= run(s, "view.freezePanes", json!({"cell": "A6"}));
    ok &= run(s, "selection.set", json!({"cell": "B2"}));
    ok
}

/// A loan: change the three inputs and the payment and a yearly balance schedule follow.
fn loan(s: &mut Session, lang: Lang) -> bool {
    let mut ok = heading(
        s,
        lang.pick("Loan calculator", "Calculadora de préstamo"),
        lang.pick("Change the three blue cells", "Cambia las tres celdas azules"),
    );
    let labels = match lang {
        Lang::En => [["Amount"], ["Yearly interest rate"], ["Years"], ["Monthly payment"], ["Total paid"], ["Total interest"]],
        Lang::Es => [["Importe"], ["Tipo de interés anual"], ["Años"], ["Cuota mensual"], ["Total pagado"], ["Intereses totales"]],
    };
    ok &= run(s, "range.setValues", json!({"range": "B5", "values": labels}));
    ok &= run(s, "cell.set", json!({"cell": "C5", "input": "180000"}));
    ok &= run(s, "cell.set", json!({"cell": "C6", "input": "0.035"}));
    ok &= run(s, "cell.set", json!({"cell": "C7", "input": "25"}));
    ok &= run(s, "cell.set", json!({"cell": "C8", "input": "=C5*(C6/12)/(1-(1+C6/12)^(-C7*12))"}));
    ok &= run(s, "cell.set", json!({"cell": "C9", "input": "=C8*C7*12"}));
    ok &= run(s, "cell.set", json!({"cell": "C10", "input": "=C9-C5"}));
    ok &= run(s, "home.numberFormat", json!({"range": "C5", "code": "#,##0.00"}));
    ok &= run(s, "home.numberFormat", json!({"range": "C6", "code": "0.00%"}));
    ok &= run(s, "home.numberFormat", json!({"range": "C8:C10", "code": "#,##0.00"}));
    ok &= run(s, "home.fontColor", json!({"range": "C5:C7", "color": "#1F5FBF"}));
    ok &= run(s, "home.fillColor", json!({"range": "C5:C7", "color": "#EAF1FB"}));
    ok &= run(s, "home.bold", json!({"range": "B8:C8", "on": true}));
    let head = match lang {
        Lang::En => [["Year", "Opening balance", "Interest", "Principal", "Closing balance"]],
        Lang::Es => [["Año", "Saldo inicial", "Intereses", "Amortizado", "Saldo final"]],
    };
    ok &= run(s, "range.setValues", json!({"range": "B13", "values": head}));
    ok &= run(s, "home.cellStyle", json!({"range": "B13:F13", "name": "Heading 3"}));
    for (i, r) in (14..=23).enumerate() {
        ok &= run(s, "cell.set", json!({"cell": format!("B{r}"), "input": (i + 1).to_string()}));
        let opening = if i == 0 { "=$C$5".to_string() } else { format!("=F{}", r - 1) };
        ok &= run(s, "cell.set", json!({"cell": format!("C{r}"), "input": opening}));
        // The balance after twelve more payments, then what that paid off and what it cost.
        let growth = "(1+$C$6/12)^12";
        ok &= run(s, "cell.set", json!({"cell": format!("F{r}"), "input": format!("=MAX(0,C{r}*{growth}-$C$8*({growth}-1)/($C$6/12))")}));
        ok &= run(s, "cell.set", json!({"cell": format!("E{r}"), "input": format!("=C{r}-F{r}")}));
        ok &= run(s, "cell.set", json!({"cell": format!("D{r}"), "input": format!("=$C$8*12-E{r}")}));
    }
    ok &= run(s, "home.numberFormat", json!({"range": "C14:F23", "code": "#,##0"}));
    ok &= run(s, "home.columnWidth", json!({"cols": "A:A", "width": 20}));
    ok &= run(s, "home.columnWidth", json!({"cols": "B:B", "width": 150}));
    ok &= run(s, "home.columnWidth", json!({"cols": "C:F", "width": 110}));
    ok &= run(s, "selection.set", json!({"cell": "B2"}));
    ok
}

/// Stock on hand: value per line and a status that says when to reorder.
fn inventory(s: &mut Session, lang: Lang) -> bool {
    let (reorder, ok_text) = (lang.pick("Reorder", "Reponer"), lang.pick("OK", "Correcto"));
    let mut ok = heading(
        s,
        lang.pick("Stock inventory", "Inventario de existencias"),
        lang.pick("A line turns red when stock falls to its reorder level", "Una línea se pone roja cuando el stock llega a su mínimo"),
    );
    let rows = match lang {
        Lang::En => json!([
            ["Item", "SKU", "In stock", "Reorder at", "Unit cost", "Value", "Status"],
            ["Watercolour set", "WC-12", 14, 6, 18.5],
            ["Sketchbook A5", "SK-A5", 40, 20, 4.2],
            ["Brush, round 6", "BR-06", 5, 8, 3.1],
            ["Ink, black 30 ml", "IN-BK", 22, 10, 5.75],
            ["Cotton paper pad", "CP-30", 8, 8, 9.9],
            ["Pencil, HB (12)", "PN-HB", 31, 12, 2.4],
        ]),
        Lang::Es => json!([
            ["Artículo", "Código", "Existencias", "Mínimo", "Coste unitario", "Valor", "Estado"],
            ["Set de acuarelas", "WC-12", 14, 6, 18.5],
            ["Cuaderno A5", "SK-A5", 40, 20, 4.2],
            ["Pincel redondo n.º 6", "BR-06", 5, 8, 3.1],
            ["Tinta negra 30 ml", "IN-BK", 22, 10, 5.75],
            ["Bloc de papel de algodón", "CP-30", 8, 8, 9.9],
            ["Lápiz HB (12)", "PN-HB", 31, 12, 2.4],
        ]),
    };
    ok &= run(s, "range.setValues", json!({"range": "B5", "values": rows}));
    for r in 6..=11 {
        ok &= run(s, "cell.set", json!({"cell": format!("G{r}"), "input": format!("=D{r}*F{r}")}));
        ok &= run(s, "cell.set", json!({"cell": format!("H{r}"), "input": format!("=IF(D{r}<=E{r},\"{reorder}\",\"{ok_text}\")")}));
    }
    ok &= run(s, "range.setValues", json!({"range": "F13", "values": [[lang.pick("Total value", "Valor total")]]}));
    ok &= run(s, "cell.set", json!({"cell": "G13", "input": "=SUM(G6:G11)"}));
    ok &= run(s, "home.bold", json!({"range": "F13:G13", "on": true}));
    ok &= run(s, "home.cellStyle", json!({"range": "B5:H5", "name": "Heading 3"}));
    ok &= run(s, "home.numberFormat", json!({"range": "F6:G13", "code": "#,##0.00"}));
    ok &= run(
        s,
        "home.conditionalFormat",
        json!({"range": "H6:H11", "rule": {"type": "cellIs", "operator": "equal", "value": format!("\"{reorder}\""), "preset": "lightRedFillDarkRedText"}}),
    );
    ok &= run(s, "home.columnWidth", json!({"cols": "A:A", "width": 20}));
    ok &= run(s, "home.columnWidth", json!({"cols": "B:B", "width": 170}));
    ok &= run(s, "home.columnWidth", json!({"cols": "C:H", "width": 96}));
    ok &= run(s, "view.freezePanes", json!({"cell": "A6"}));
    ok &= run(s, "selection.set", json!({"cell": "B2"}));
    ok
}

/// A week of hours per project, with totals across and down and what they come to at a rate.
fn timesheet(s: &mut Session, lang: Lang) -> bool {
    let mut ok = heading(
        s,
        lang.pick("Weekly timesheet", "Hoja de horas semanal"),
        lang.pick("Type the hours; the totals and the amount follow", "Escribe las horas; los totales y el importe se calculan solos"),
    );
    let rows = match lang {
        Lang::En => json!([
            ["Project", "Mon", "Tue", "Wed", "Thu", "Fri", "Total"],
            ["Book cover", 3, 4, 2, 0, 0],
            ["Chapter openers", 2, 1, 4, 5, 3],
            ["Revisions", 0, 1, 0, 1, 2],
            ["Admin", 1, 1, 1, 1, 1],
        ]),
        Lang::Es => json!([
            ["Proyecto", "Lun", "Mar", "Mié", "Jue", "Vie", "Total"],
            ["Portada del libro", 3, 4, 2, 0, 0],
            ["Capitulares", 2, 1, 4, 5, 3],
            ["Revisiones", 0, 1, 0, 1, 2],
            ["Gestión", 1, 1, 1, 1, 1],
        ]),
    };
    ok &= run(s, "range.setValues", json!({"range": "B5", "values": rows}));
    for r in 6..=9 {
        ok &= run(s, "cell.set", json!({"cell": format!("H{r}"), "input": format!("=SUM(C{r}:G{r})")}));
    }
    ok &= run(s, "range.setValues", json!({"range": "B10", "values": [[lang.pick("Total", "Total")]]}));
    for c in ["C", "D", "E", "F", "G", "H"] {
        ok &= run(s, "cell.set", json!({"cell": format!("{c}10"), "input": format!("=SUM({c}6:{c}9)")}));
    }
    ok &= run(
        s,
        "range.setValues",
        json!({"range": "B12", "values": [[lang.pick("Hourly rate", "Precio por hora")], [lang.pick("Amount", "Importe")]]}),
    );
    ok &= run(s, "cell.set", json!({"cell": "C12", "input": "45"}));
    ok &= run(s, "cell.set", json!({"cell": "C13", "input": "=H10*C12"}));
    ok &= run(s, "home.numberFormat", json!({"range": "C12:C13", "code": "#,##0.00"}));
    ok &= run(s, "home.fontColor", json!({"range": "C12", "color": "#1F5FBF"}));
    ok &= run(s, "home.fillColor", json!({"range": "C12", "color": "#EAF1FB"}));
    ok &= run(s, "home.cellStyle", json!({"range": "B5:H5", "name": "Heading 3"}));
    ok &= run(s, "home.cellStyle", json!({"range": "B10:H10", "name": "Total"}));
    ok &= run(s, "home.bold", json!({"range": "B13:C13", "on": true}));
    ok &= run(s, "home.conditionalFormat", json!({"range": "C6:G9", "rule": {"type": "colorScale", "colors": ["#FFFFFF", "#63BE7B"]}}));
    ok &= run(s, "home.columnWidth", json!({"cols": "A:A", "width": 20}));
    ok &= run(s, "home.columnWidth", json!({"cols": "B:B", "width": 150}));
    ok &= run(s, "home.columnWidth", json!({"cols": "C:H", "width": 72}));
    ok &= run(s, "selection.set", json!({"cell": "B2"}));
    ok
}

/// A savings goal: three inputs, how many months it takes, and a year of progress toward it.
fn savings(s: &mut Session, lang: Lang) -> bool {
    let mut ok = heading(
        s,
        lang.pick("Savings goal", "Objetivo de ahorro"),
        lang.pick("Change the three blue cells and the plan follows", "Cambia las tres celdas azules y el plan se ajusta"),
    );
    let labels = match lang {
        Lang::En => [["Goal"], ["Saved so far"], ["Saved each month"], ["Months to go"], ["Progress"]],
        Lang::Es => [["Objetivo"], ["Ahorrado hasta ahora"], ["Ahorro mensual"], ["Meses que faltan"], ["Avance"]],
    };
    ok &= run(s, "range.setValues", json!({"range": "B5", "values": labels}));
    ok &= run(s, "cell.set", json!({"cell": "C5", "input": "12000"}));
    ok &= run(s, "cell.set", json!({"cell": "C6", "input": "2500"}));
    ok &= run(s, "cell.set", json!({"cell": "C7", "input": "400"}));
    ok &= run(s, "cell.set", json!({"cell": "C8", "input": "=IF(C7>0,MAX(0,ROUNDUP((C5-C6)/C7,0)),0)"}));
    ok &= run(s, "cell.set", json!({"cell": "C9", "input": "=IF(C5>0,MIN(1,C6/C5),0)"}));
    ok &= run(s, "home.numberFormat", json!({"range": "C5:C7", "code": "#,##0"}));
    ok &= run(s, "home.numberFormat", json!({"range": "C9", "code": "0%"}));
    ok &= run(s, "home.fontColor", json!({"range": "C5:C7", "color": "#1F5FBF"}));
    ok &= run(s, "home.fillColor", json!({"range": "C5:C7", "color": "#EAF1FB"}));
    ok &= run(s, "home.bold", json!({"range": "B8:C8", "on": true}));
    let head = match lang {
        Lang::En => [["Month", "Saved", "Of the goal"]],
        Lang::Es => [["Mes", "Ahorrado", "Del objetivo"]],
    };
    ok &= run(s, "range.setValues", json!({"range": "B12", "values": head}));
    ok &= run(s, "home.cellStyle", json!({"range": "B12:D12", "name": "Heading 3"}));
    for (i, r) in (13..=24).enumerate() {
        ok &= run(s, "cell.set", json!({"cell": format!("B{r}"), "input": (i + 1).to_string()}));
        ok &= run(s, "cell.set", json!({"cell": format!("C{r}"), "input": format!("=MIN($C$5,$C$6+$C$7*B{r})")}));
        ok &= run(s, "cell.set", json!({"cell": format!("D{r}"), "input": format!("=IF($C$5>0,C{r}/$C$5,0)")}));
    }
    ok &= run(s, "home.numberFormat", json!({"range": "C13:C24", "code": "#,##0"}));
    ok &= run(s, "home.numberFormat", json!({"range": "D13:D24", "code": "0%"}));
    ok &= run(s, "home.conditionalFormat", json!({"range": "D13:D24", "rule": {"type": "dataBar", "color": "#63BE7B"}}));
    ok &= run(s, "home.columnWidth", json!({"cols": "A:A", "width": 20}));
    ok &= run(s, "home.columnWidth", json!({"cols": "B:B", "width": 170}));
    ok &= run(s, "home.columnWidth", json!({"cols": "C:D", "width": 120}));
    ok &= run(s, "selection.set", json!({"cell": "B2"}));
    ok
}

/// An expense report: dated lines by category, and what each category and the whole trip came to.
fn expenses(s: &mut Session, lang: Lang) -> bool {
    let (travel, meals, lodging) = (lang.pick("Travel", "Viajes"), lang.pick("Meals", "Comidas"), lang.pick("Lodging", "Alojamiento"));
    let mut ok = heading(
        s,
        lang.pick("Expense report", "Informe de gastos"),
        lang.pick("Add a line under the last one; the totals follow", "Añade una línea bajo la última; los totales se ajustan"),
    );
    let rows = match lang {
        Lang::En => json!([
            ["Date", "Category", "Description", "Amount"],
            ["2026-10-12", travel, "Train to the book fair", 64.5],
            ["2026-10-12", meals, "Lunch with the editor", 28.9],
            ["2026-10-13", lodging, "Two nights, guesthouse", 140],
            ["2026-10-14", travel, "Taxi to the station", 18],
            ["2026-10-14", meals, "Dinner on the way home", 22.4],
        ]),
        Lang::Es => json!([
            ["Fecha", "Categoría", "Descripción", "Importe"],
            ["2026-10-12", travel, "Tren a la feria del libro", 64.5],
            ["2026-10-12", meals, "Comida con la editora", 28.9],
            ["2026-10-13", lodging, "Dos noches, casa de huéspedes", 140],
            ["2026-10-14", travel, "Taxi a la estación", 18],
            ["2026-10-14", meals, "Cena de camino a casa", 22.4],
        ]),
    };
    ok &= run(s, "range.setValues", json!({"range": "B5", "values": rows}));
    ok &= run(s, "home.cellStyle", json!({"range": "B5:E5", "name": "Heading 3"}));
    ok &= run(s, "home.numberFormat", json!({"range": "E6:E10", "code": "#,##0.00"}));
    ok &= run(
        s,
        "range.setValues",
        json!({"range": "B13", "values": [[lang.pick("By category", "Por categoría")], [travel], [meals], [lodging], [lang.pick("Total", "Total")]]}),
    );
    ok &= run(s, "home.cellStyle", json!({"range": "B13:C13", "name": "Heading 3"}));
    for r in 14..=16 {
        ok &= run(s, "cell.set", json!({"cell": format!("C{r}"), "input": format!("=SUMIF($C$6:$C$10,B{r},$E$6:$E$10)")}));
    }
    ok &= run(s, "cell.set", json!({"cell": "C17", "input": "=SUM(C14:C16)"}));
    ok &= run(s, "home.cellStyle", json!({"range": "B17:C17", "name": "Total"}));
    ok &= run(s, "home.numberFormat", json!({"range": "C14:C17", "code": "#,##0.00"}));
    ok &= run(
        s,
        "data.validation",
        json!({"range": "C6:C10", "type": "list", "formula1": format!("\"{travel},{meals},{lodging}\""), "dropdown": true}),
    );
    ok &= run(s, "home.columnWidth", json!({"cols": "A:A", "width": 20}));
    ok &= run(s, "home.columnWidth", json!({"cols": "B:C", "width": 120}));
    ok &= run(s, "home.columnWidth", json!({"cols": "D:D", "width": 230}));
    ok &= run(s, "home.columnWidth", json!({"cols": "E:E", "width": 100}));
    ok &= run(s, "selection.set", json!({"cell": "B2"}));
    ok
}

/// A week at a glance: three parts of the day for each weekday, the weekend tinted.
fn planner(s: &mut Session, lang: Lang) -> bool {
    let mut ok = heading(
        s,
        lang.pick("Weekly planner", "Planificador semanal"),
        lang.pick("One line for each part of the day", "Una línea para cada parte del día"),
    );
    let rows = match lang {
        Lang::En => json!([
            ["Day", "Morning", "Afternoon", "Evening"],
            ["Monday", "Plan the week", "Studio: sketches", "Swim"],
            ["Tuesday", "Emails and invoices", "Studio: inks", "Cook for friends"],
            ["Wednesday", "Library", "Client call at 3", ""],
            ["Thursday", "Studio: colour", "Open studio evening setup", "Open studio"],
            ["Friday", "Revisions", "Send proofs", "Film night"],
            ["Saturday", "Market", "Garden", ""],
            ["Sunday", "Long breakfast", "Walk by the harbour", "Read"],
        ]),
        Lang::Es => json!([
            ["Día", "Mañana", "Tarde", "Noche"],
            ["Lunes", "Planificar la semana", "Taller: bocetos", "Natación"],
            ["Martes", "Correos y facturas", "Taller: tintas", "Cena con amigos"],
            ["Miércoles", "Biblioteca", "Llamada con el cliente a las 15:00", ""],
            ["Jueves", "Taller: color", "Preparar la tarde abierta", "Taller abierto"],
            ["Viernes", "Revisiones", "Enviar pruebas", "Noche de cine"],
            ["Sábado", "Mercado", "Jardín", ""],
            ["Domingo", "Desayuno largo", "Paseo por el puerto", "Leer"],
        ]),
    };
    ok &= run(s, "range.setValues", json!({"range": "B5", "values": rows}));
    ok &= run(s, "home.cellStyle", json!({"range": "B5:E5", "name": "Heading 3"}));
    ok &= run(s, "home.bold", json!({"range": "B6:B12", "on": true}));
    ok &= run(s, "home.fillColor", json!({"range": "B11:E12", "color": "#EEF4F5"}));
    ok &= run(s, "home.columnWidth", json!({"cols": "A:A", "width": 20}));
    ok &= run(s, "home.columnWidth", json!({"cols": "B:B", "width": 110}));
    ok &= run(s, "home.columnWidth", json!({"cols": "C:E", "width": 200}));
    ok &= run(s, "selection.set", json!({"cell": "B2"}));
    ok
}

/// A blank sheet (for callers that need one without a session).
pub fn blank_sheet(name: &str) -> Sheet {
    Sheet::new(name)
}
