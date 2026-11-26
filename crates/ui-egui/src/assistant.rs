//! The AI assistant: Celdas's side of the Office suite's shared chat (`nodalia-assistant`, from
//! nodalia-design `packages/assistant`).
//!
//! The component draws the button, the panel and the Preferences section the same way in Letra,
//! Lámina and here, keeps the person's provider, model and key, and runs the loop with the model.
//! This file is everything the component cannot know: what a workbook is, which tools the model
//! gets to work on it, and how a turn is undone. The tools sit on the same command registry that
//! scripts and MCP use, so the assistant can do what a person can do from the ribbon and nothing
//! else: it cannot open, save, close or print files, and it does not run macros.
//!
//! The person's own messages are the only instructions. The cells are data, and every change
//! waits for Apply unless they chose to let the assistant work as it goes (Preferences). A whole
//! turn undoes in one click.

use std::sync::Arc;

use celdas_engine::core::{CellRef, RangeRef};
use egui::Ui;
use nodalia_assistant::{Assistant, Host, Lang, ToolSpec};
use nodalia_theme::office::Look;
use serde_json::{Value, json};

use crate::SheetApp;
use crate::theme::{self, Tokens};

/// Rows and columns of the active sheet the snapshot shows; the model reads the rest with
/// `read_range`.
const PREVIEW_ROWS: u32 = 60;
const PREVIEW_COLS: u32 = 20;
/// Characters of a text cell the snapshot keeps.
const LONG_CELL: usize = 200;
/// Most commands or functions one answer lists.
const MAX_LISTED: usize = 60;
/// Longest text the approval card shows for one argument.
const PREVIEW: usize = 80;

/// Command families the assistant never runs: files and windows are the person's to open, save,
/// close and print, `ui.*` drives the window rather than the workbook, and a macro is code that
/// the person runs themselves.
const FORBIDDEN: [&str; 8] = ["file.", "ui.", "app.", "window.", "print", "macro.run", "macro.debug", "macro.record"];

/// The look the shared component draws with: Celdas's palette and its bold family.
fn look(t: &Tokens) -> Look {
    Look::new(theme::palette_of(t)).with_bold(egui::FontFamily::Name(theme::UI_BOLD.into()))
}

/// The language the component speaks: Spanish when the interface is Spanish, English otherwise.
fn lang(ctx: &egui::Context) -> Lang {
    Lang::from_code(crate::i18n::current(ctx).code())
}

impl SheetApp {
    /// The assistant, made the first time it is asked for. It reads the saved choices and starts
    /// asking the keychain, which nothing should do before a window is on screen.
    pub(crate) fn assistant_mut(&mut self, ctx: &egui::Context) -> &mut Assistant {
        let ctx = ctx.clone();
        self.assistant.get_or_insert_with(|| Assistant::new(Arc::new(move || ctx.request_repaint())))
    }

    /// Runs `f` with the assistant and a host over the active workbook. The assistant lives in the
    /// app and the host borrows the app, so it is taken out for the duration.
    fn with_assistant<R>(&mut self, ctx: &egui::Context, f: impl FnOnce(&mut Assistant, &mut CeldasHost<'_>) -> R) -> R {
        self.assistant_mut(ctx);
        let lang = lang(ctx);
        let mut assistant = self.assistant.take().unwrap_or_else(|| Assistant::new(Arc::new(|| {})));
        assistant.lang = lang;
        let result = f(&mut assistant, &mut CeldasHost { app: self, lang });
        self.assistant = Some(assistant);
        result
    }

    /// Opens or closes the assistant's chat (`ui.assistant`, and the snapshot example). `None`
    /// switches it. Returns whether it is open.
    pub fn show_assistant(&mut self, ctx: &egui::Context, open: Option<bool>) -> bool {
        let a = self.assistant_mut(ctx);
        a.open = open.unwrap_or(!a.open);
        a.open
    }

    /// Runs the model's pending tool calls and moves its answer along. Called once a frame.
    pub(crate) fn poll_assistant(&mut self, ctx: &egui::Context) {
        if self.assistant.is_none() {
            return;
        }
        if self.with_assistant(ctx, |a, host| a.poll(host)) {
            ctx.request_repaint();
        }
    }
}

/// Whether the title bar offers the assistant: the person did not turn it off.
pub(crate) fn offered(app: &mut SheetApp, ctx: &egui::Context) -> bool {
    app.assistant_mut(ctx).settings.enabled
}

/// The title bar's button: the chosen provider's mark, which opens and closes the chat.
pub(crate) fn button(app: &mut SheetApp, ui: &mut Ui) {
    let look = look(&Tokens::get(ui.ctx()));
    let lang = lang(ui.ctx());
    let a = app.assistant_mut(ui.ctx());
    a.lang = lang;
    let label = lang.t("Assistant");
    let r = nodalia_assistant::ui::button(ui, &look, a);
    r.widget_info(|| egui::WidgetInfo::labeled(egui::Role::Button, true, label));
}

/// The chat panel, on the right edge, drawn before the sheet so it keeps the edge.
pub(crate) fn panel(app: &mut SheetApp, ui: &mut Ui) {
    if !app.assistant.as_ref().is_some_and(|a| a.open) {
        return;
    }
    let ctx = ui.ctx().clone();
    let t = Tokens::get(&ctx);
    let look = look(&t);
    let gap = theme::ISLAND_GAP;
    let mut open_preferences = false;
    egui::Panel::right("assistant_panel")
        .exact_size(nodalia_assistant::ui::PANEL_WIDTH + gap)
        .show_separator_line(false)
        .frame(egui::Frame::NONE.fill(t.window).inner_margin(egui::Margin { left: 0, right: gap as i8, top: 0, bottom: gap as i8 }))
        .show(ui, |ui| {
            open_preferences = app.with_assistant(&ctx, |a, host| nodalia_assistant::ui::panel(ui, &look, a, host)).open_preferences;
        });
    if open_preferences {
        app.open_backstage(crate::backstage::Page::Options);
    }
}

/// The Preferences section.
pub(crate) fn preferences(app: &mut SheetApp, ui: &mut Ui) {
    let look = look(&Tokens::get(ui.ctx()));
    let a = app.assistant_mut(ui.ctx());
    a.lang = lang(ui.ctx());
    nodalia_assistant::ui::preferences(ui, &look, a);
}

/// What the assistant works through: the active workbook, by the commands the registry has.
pub(crate) struct CeldasHost<'a> {
    app: &'a mut SheetApp,
    lang: Lang,
}

fn tool(name: &str, description: &str, schema: Value, read_only: bool) -> ToolSpec {
    ToolSpec { name: name.into(), description: description.into(), schema, read_only }
}

fn object(properties: Value, required: &[&str]) -> Value {
    json!({"type": "object", "properties": properties, "required": required})
}

fn text(description: &str) -> Value {
    json!({"type": "string", "description": description})
}

fn flag(description: &str) -> Value {
    json!({"type": "boolean", "description": description})
}

fn forbidden(command: &str) -> bool {
    FORBIDDEN.iter().any(|f| command.starts_with(f))
}

/// The first `n` characters of `s`, with an ellipsis when it was longer.
fn clip(s: &str, n: usize) -> String {
    let mut out: String = s.chars().take(n).collect();
    if s.chars().nth(n).is_some() {
        out.push('…');
    }
    out
}

fn str_arg<'v>(input: &'v Value, key: &str) -> Result<&'v str, String> {
    input.get(key).and_then(Value::as_str).ok_or_else(|| format!("`{key}` is missing; it must be a string"))
}

/// Copies the listed keys (when present and not null) into a params object.
fn pick(input: &Value, keys: &[&str]) -> Value {
    Value::Object(keys.iter().filter_map(|k| input.get(*k).filter(|v| !v.is_null()).map(|v| (k.to_string(), v.clone()))).collect())
}

/// Shortens every long text in rows of cells.
fn clip_cells(rows: &mut Value) {
    let Some(rows) = rows.as_array_mut() else { return };
    for row in rows {
        let Some(cells) = row.as_array_mut() else { continue };
        for cell in cells {
            if let Some(t) = cell.as_str() {
                *cell = json!(clip(t, LONG_CELL));
            }
        }
    }
}

/// The steps of `format_range`: one command per property given.
fn format_steps(input: &Value) -> Result<Vec<(&'static str, Value)>, String> {
    let range = str_arg(input, "range")?;
    let mut steps: Vec<(&'static str, Value)> = Vec::new();
    let mut add = |id: &'static str, extra: Value| {
        let mut params = json!({"range": range});
        if let (Some(p), Some(e)) = (params.as_object_mut(), extra.as_object()) {
            p.extend(e.clone());
        }
        steps.push((id, params));
    };
    for (key, id) in [("bold", "home.bold"), ("italic", "home.italic"), ("underline", "home.underline"), ("wrapText", "home.wrapText")] {
        if let Some(on) = input.get(key).and_then(Value::as_bool) {
            add(id, json!({"on": on}));
        }
    }
    if let Some(v) = input.get("fontName").and_then(Value::as_str) {
        add("home.fontName", json!({"name": v}));
    }
    if let Some(v) = input.get("fontSize").and_then(Value::as_f64) {
        add("home.fontSize", json!({"size": v}));
    }
    if let Some(v) = input.get("fontColor").and_then(Value::as_str) {
        add("home.fontColor", json!({"color": v}));
    }
    if let Some(v) = input.get("fillColor").and_then(Value::as_str) {
        add("home.fillColor", json!({"color": v}));
    }
    if let Some(v) = input.get("numberFormat").and_then(Value::as_str) {
        add("home.numberFormat", json!({"format": v}));
    }
    if let Some(v) = input.get("horizontalAlign").and_then(Value::as_str) {
        let id = match v {
            "left" => "home.alignLeft",
            "center" => "home.alignCenter",
            "right" => "home.alignRight",
            "justify" => "home.justify",
            other => return Err(format!("horizontalAlign `{other}` must be left, center, right or justify")),
        };
        add(id, json!({}));
    }
    match input.get("borders") {
        None | Some(Value::Null) => {}
        Some(Value::String(preset)) => add("home.borders", json!({"preset": preset})),
        Some(Value::Object(o)) => add("home.borders", Value::Object(o.clone())),
        Some(_) => return Err("`borders` must be a preset name or {preset, style?, color?}".into()),
    }
    if steps.is_empty() {
        return Err("nothing to apply: give at least one of bold, italic, underline, fontName, fontSize, fontColor, fillColor, numberFormat, horizontalAlign, wrapText, borders".into());
    }
    Ok(steps)
}

/// Sort keys as the `data.sort` command takes them: `"B"`, `"B desc"` or `{column, order}`.
fn sort_keys(input: &Value) -> Result<Value, String> {
    let keys = input.get("keys").and_then(Value::as_array).ok_or("`keys` must be an array of columns")?;
    if keys.is_empty() {
        return Err("`keys` needs at least one sort level".into());
    }
    let mut out = Vec::with_capacity(keys.len());
    for key in keys {
        out.push(match key {
            Value::String(c) => {
                let mut words = c.split_whitespace();
                let column = words.next().unwrap_or_default();
                let order = match words.next().map(str::to_ascii_lowercase).as_deref() {
                    Some("desc" | "descending") => "desc",
                    _ => "asc",
                };
                json!({"column": column, "order": order})
            }
            Value::Object(o) => {
                let mut o = o.clone();
                o.entry("order").or_insert(json!("asc"));
                Value::Object(o)
            }
            _ => return Err("each sort key is a column letter or {column, order}".into()),
        });
    }
    Ok(Value::Array(out))
}

impl CeldasHost<'_> {
    fn run(&mut self, command: &str, params: Value) -> Result<Value, String> {
        self.app.run(command, params)
    }

    /// A cell the person is still typing into belongs to the workbook before the assistant changes
    /// it, so the draft is entered first and never lost or overwritten.
    fn edit(&mut self, command: &str, params: Value) -> Result<Value, String> {
        self.app.commit_edit(0, 0, false, false);
        self.run(command, params)
    }

    fn list_commands(&mut self, query: &str) -> Value {
        let q = query.to_lowercase();
        let found: Vec<Value> = self
            .app
            .session
            .commands()
            .into_iter()
            .filter(|c| !forbidden(c.id))
            .filter(|c| {
                q.is_empty()
                    || c.id.to_lowercase().contains(&q)
                    || c.label.to_lowercase().contains(&q)
                    || c.menu.iter().any(|m| m.to_lowercase().contains(&q))
            })
            .take(MAX_LISTED)
            .map(|c| json!({"id": c.id, "label": c.label, "where": c.menu.join(" › "), "params": c.params, "enabled": c.enabled}))
            .collect();
        json!({"commands": found, "limit": MAX_LISTED})
    }

    fn list_functions(&mut self, input: &Value) -> Result<Value, String> {
        let mut found = self.run("formulas.functions", pick(input, &["search", "category"]))?;
        let total = found.as_array().map(Vec::len);
        if let (Some(list), Some(total)) = (found.as_array_mut(), total)
            && total > MAX_LISTED
        {
            list.truncate(MAX_LISTED);
            return Ok(json!({"functions": found, "shown": MAX_LISTED, "total": total, "note": "Narrow the search to see the rest."}));
        }
        Ok(json!({"functions": found}))
    }

    /// The top-left corner of the active sheet's data, with formulas as the person typed them.
    fn preview(&mut self) -> Value {
        let used = self.app.session.active().and_then(|d| d.wb.active()).and_then(|s| s.used_range());
        let Some(used) = used else { return Value::Null };
        let end = CellRef::new(
            used.end.row.min(used.start.row.saturating_add(PREVIEW_ROWS - 1)),
            used.end.col.min(used.start.col.saturating_add(PREVIEW_COLS - 1)),
        );
        let shown = RangeRef::new(used.start, end);
        let Ok(mut read) = self.run("sheet.read", json!({"range": shown.a1(), "formulas": true})) else { return Value::Null };
        if let Some(values) = read.get_mut("values") {
            clip_cells(values);
        }
        json!({
            "range": shown.a1(),
            "usedRange": used.a1(),
            "values": read.get("values"),
            "note": (shown.a1() != used.a1()).then_some("Only the top-left corner of the sheet is shown. Read the rest with read_range."),
        })
    }
}

const GUIDE: &str = "Celdas is a spreadsheet. A workbook has sheets of cells addressed in A1 style (\"B7\", \"A1:D10\", \"Sheet2!A1:C3\"); a cell holds a number, text, a boolean, an error or a formula that starts with =. Formulas use English function names and commas between arguments.\n\
\n\
The snapshot gives the workbook's sheets (used range, tables, charts, filters), the selection and a preview of the active sheet's top-left corner with formulas as typed. Where the person is working is the selection. Read anything else with read_range and read single cells with get_cell.\n\
\n\
Prefer formulas to computed numbers so the sheet stays live: write =SUM(B2:B9), not 123. Check a formula with evaluate_formula before you enter it. Write many cells in one call with write_range (a 2-D array starting at the top-left cell) rather than cell by cell. A chart's first row and column are its series names and categories. Format a range with format_range, build a table with create_table, order rows with sort_range and narrow them with filter.\n\
\n\
Celdas has hundreds of commands, one for each ribbon button. Look one up with list_commands (search by what it does, like \"conditional\", \"pivot\", \"freeze\") and read its params before running it with run_command. You cannot open, save, close or print files, and you do not run macros.";

impl Host for CeldasHost<'_> {
    fn app_name(&self) -> &str {
        "Celdas"
    }

    fn guide(&self) -> String {
        GUIDE.into()
    }

    fn context(&mut self) -> Result<Value, String> {
        let mut inspected = self.run("document.inspect", json!({}))?;
        if let Some(o) = inspected.as_object_mut() {
            // The person's own undo labels say nothing about the data.
            o.remove("undo");
            o.remove("redo");
            o.insert("preview".into(), self.preview());
        }
        Ok(inspected)
    }

    fn tools(&self) -> Vec<ToolSpec> {
        let range = "A1-style range, e.g. \"A1:D10\", \"B:B\" or \"Sheet2!A1:C3\"";
        vec![
            tool(
                "read_range",
                "Read cell values from a range as rows. Numbers, booleans and text come back as JSON values, empty cells as null and errors as {\"error\":\"#DIV/0!\"}. Without `range`, reads the sheet's used range. formulas=true returns formula text for formula cells; formatted=true returns the text as displayed.",
                object(
                    json!({"range": text(range), "sheet": text("Sheet name (default: the active sheet)"), "formulas": flag("Return formulas instead of values"), "formatted": flag("Return the displayed text")}),
                    &[],
                ),
                true,
            ),
            tool(
                "get_cell",
                "Details of one cell: computed value and type, formula, displayed text, style, comment and link.",
                object(json!({"cell": text("Cell, e.g. \"B7\" or \"Sheet2!C3\"")}), &["cell"]),
                true,
            ),
            tool(
                "evaluate_formula",
                "Evaluate a formula WITHOUT changing any cell, e.g. \"=SUM(B2:B10)/COUNT(B2:B10)\". `cell` sets the position for relative references. Use English function names and commas.",
                object(
                    json!({"formula": text("Formula text, with or without the leading ="), "cell": text("Evaluate as if in this cell (default: the active cell)")}),
                    &["formula"],
                ),
                true,
            ),
            tool(
                "list_functions",
                "List worksheet functions: name, category, signature and description. Filter with `search` or `category`. Returns up to 60.",
                object(json!({"search": text("Part of the name or description"), "category": text("A category, e.g. \"Lookup & Reference\"")}), &[]),
                true,
            ),
            tool(
                "list_commands",
                "Find Celdas commands by what they do. Searches each command's id, label and ribbon path and says whether it can run now. Returns up to 60.",
                object(json!({"query": text("Words to look for, e.g. \"pivot\", \"conditional\", \"freeze\"")}), &[]),
                true,
            ),
            tool(
                "select",
                "Set the selection, which is what run_command acts on when it has no `range`. Several areas with commas: \"A1:B2,D4\". A \"Sheet2!A1\" prefix switches sheets. Does not change any data.",
                object(json!({"range": text(range), "active": text("The active cell inside the selection")}), &["range"]),
                true,
            ),
            tool(
                "write_range",
                "Write a 2-D array of values starting at the top-left cell of `range`. Numbers, booleans, null (clears) and text are accepted; text starting with = is a formula, other text is read like typed input (\"12%\", \"1/2/2024\"). One undo step.",
                object(
                    json!({
                        "range": text("Top-left cell, e.g. \"A1\" or \"Sheet2!B3\""),
                        "values": {"type": "array", "description": "Rows of cell values", "items": {"type": "array", "items": {"type": ["string", "number", "boolean", "null"]}}},
                        "parse": flag("Read text like typed input (default true); false stores it as the text it is"),
                    }),
                    &["range", "values"],
                ),
                false,
            ),
            tool(
                "set_cell",
                "Enter one cell as if typed: a number, text or a formula starting with =. Dependent cells recalculate.",
                object(
                    json!({"cell": text("Cell, e.g. \"B7\""), "input": {"type": ["string", "number", "boolean"], "description": "Number, text or =formula"}}),
                    &["cell", "input"],
                ),
                false,
            ),
            tool(
                "format_range",
                "Apply several formats to a range in one call: bold, italic, underline, font, colours (\"#RRGGBB\"), number format (a name like Currency, Percentage, Short Date or a code like \"#,##0.00\"), alignment, wrap and borders.",
                object(
                    json!({
                        "range": text(range),
                        "bold": flag("Bold on or off"),
                        "italic": flag("Italic on or off"),
                        "underline": flag("Underline on or off"),
                        "fontName": text("Font family"),
                        "fontSize": {"type": "number", "minimum": 1, "maximum": 409, "description": "Size in points"},
                        "fontColor": text("Text colour \"#RRGGBB\" or \"auto\""),
                        "fillColor": text("Background colour \"#RRGGBB\" or \"none\""),
                        "numberFormat": text("A format name or code"),
                        "horizontalAlign": {"type": "string", "enum": ["left", "center", "right", "justify"]},
                        "wrapText": flag("Wrap text on or off"),
                        "borders": {"description": "A preset (all, outside, bottom, none…) or {preset, style?, color?}"},
                    }),
                    &["range"],
                ),
                false,
            ),
            tool(
                "insert_chart",
                "Insert a chart from a data range, whose first row and column are the series names and categories. Returns the chart's id.",
                object(
                    json!({
                        "range": text("Data including headers, e.g. \"A1:C6\""),
                        "type": {"type": "string", "enum": ["column", "bar", "line", "pie", "doughnut", "area", "scatter", "bubble", "radar", "combo", "histogram", "waterfall", "funnel", "treemap", "sunburst", "boxWhisker", "stock"]},
                        "subtype": {"type": "string", "enum": ["clustered", "stacked", "stacked100", "markers"]},
                        "title": text("Chart title"),
                        "at": text("Top-left cell of the chart (default: right of the data)"),
                    }),
                    &["range", "type"],
                ),
                false,
            ),
            tool(
                "create_table",
                "Turn a range into a table with banded rows and filter buttons. The first row is the header unless header=false.",
                object(
                    json!({"range": text("Range with the header row, e.g. \"A1:D20\""), "style": text("A table style, e.g. \"TableStyleMedium2\""), "header": flag("The first row holds headers (default true)"), "name": text("Table name")}),
                    &["range"],
                ),
                false,
            ),
            tool(
                "sort_range",
                "Sort a range by one or more columns, most significant first: [\"B\"], [\"B desc\"] or [{\"column\":\"B\",\"order\":\"desc\"}]. Without `range` the region around the selection is sorted.",
                object(
                    json!({"range": text(range), "keys": {"type": "array", "description": "Sort levels", "items": {}}, "header": flag("The first row is a header and stays in place")}),
                    &["keys"],
                ),
                false,
            ),
            tool(
                "filter",
                "Filter rows by a column, turning AutoFilter on when needed. Give `values` to keep, `custom` ({op, value}), `top` ({count, bottom?, percent?}) or clear=true.",
                object(
                    json!({
                        "column": text("Column letter, e.g. \"C\""),
                        "range": text("Data range (default: the region around the selection)"),
                        "values": {"type": "array", "items": {"type": ["string", "number", "boolean"]}, "description": "Values to keep"},
                        "custom": {"type": "object", "description": "{op: \"=\"|\"<>\"|\">\"|\">=\"|\"<\"|\"<=\"|\"beginsWith\"|\"contains\", value, op2?, value2?, and?}"},
                        "top": {"type": "object", "description": "{count, bottom?, percent?}"},
                        "blanks": flag("Also keep blank cells"),
                        "clear": flag("Remove this column's filter"),
                    }),
                    &["column"],
                ),
                false,
            ),
            tool(
                "run_command",
                "Run a Celdas command by id with params, as the ribbon button would on the selection, e.g. data.removeDuplicates or sheet.insert. Find ids with list_commands. File, window, print and macro-run commands are not available.",
                object(json!({"command": text("The command id"), "params": {"type": "object", "description": "The command's params"}}), &["command"]),
                false,
            ),
        ]
    }

    fn call(&mut self, tool: &str, input: &Value) -> Result<Value, String> {
        match tool {
            "read_range" => self.run("sheet.read", pick(input, &["range", "sheet", "formulas", "formatted"])),
            "get_cell" => self.run("cell.get", pick(input, &["cell"])),
            "evaluate_formula" => {
                let formula = str_arg(input, "formula")?;
                let formula = if formula.starts_with('=') { formula.to_string() } else { format!("={formula}") };
                let mut params = pick(input, &["cell"]);
                params["formula"] = json!(formula);
                let result = self.run("formulas.evaluate", params)?;
                Ok(json!({"formula": formula, "result": result}))
            }
            "list_functions" => self.list_functions(input),
            "list_commands" => Ok(self.list_commands(input.get("query").and_then(Value::as_str).unwrap_or(""))),
            "select" => self.run("selection.set", pick(input, &["range", "active"])),
            "write_range" => self.edit("range.setValues", pick(input, &["range", "values", "parse"])),
            "set_cell" => self.edit("cell.set", pick(input, &["cell", "input"])),
            "format_range" => {
                let steps = format_steps(input)?;
                let mut applied: Vec<&str> = Vec::new();
                for (id, params) in steps {
                    self.edit(id, params).map_err(|e| format!("{id}: {e} (applied before the failure: {applied:?})"))?;
                    applied.push(id);
                }
                Ok(json!({"range": str_arg(input, "range")?, "applied": applied}))
            }
            "insert_chart" => self.edit("insert.chart", pick(input, &["range", "type", "subtype", "title", "at"])),
            "create_table" => self.edit("insert.table", pick(input, &["range", "style", "header", "name"])),
            "sort_range" => {
                let mut params = pick(input, &["range", "header"]);
                params["keys"] = sort_keys(input)?;
                self.edit("data.sort", params)
            }
            "filter" => {
                if let Some(range) = input.get("range").and_then(Value::as_str) {
                    self.edit("selection.set", json!({"range": range}))?;
                }
                self.edit("data.filterBy", pick(input, &["column", "values", "custom", "top", "blanks", "clear"]))
            }
            "run_command" => {
                let command = str_arg(input, "command")?;
                if forbidden(command) {
                    return Err(format!(
                        "`{command}` is not available to the assistant: files, windows, printing and macros are the person's to handle"
                    ));
                }
                let params = input.get("params").cloned().unwrap_or_else(|| json!({}));
                self.edit(command, params)
            }
            other => Err(format!("there is no tool called `{other}`")),
        }
    }

    fn describe(&self, tool: &str, input: &Value) -> Option<String> {
        let es = self.lang == Lang::Spanish;
        let arg = |k: &str| input.get(k).and_then(Value::as_str).map(|s| clip(s, PREVIEW));
        Some(match tool {
            "write_range" => {
                let at = arg("range")?;
                let rows = input.get("values").and_then(Value::as_array)?;
                let cols = rows.iter().filter_map(Value::as_array).map(Vec::len).max().unwrap_or(0);
                if es {
                    format!("Escribir {} × {} celdas desde {at}", rows.len(), cols)
                } else {
                    format!("Write {} × {} cells from {at}", rows.len(), cols)
                }
            }
            "set_cell" => {
                let cell = arg("cell")?;
                let input = match input.get("input")? {
                    Value::String(s) => clip(s, PREVIEW),
                    other => other.to_string(),
                };
                if es { format!("Escribir «{input}» en {cell}") } else { format!("Enter “{input}” in {cell}") }
            }
            "format_range" => {
                let range = arg("range")?;
                if es { format!("Dar formato a {range}") } else { format!("Format {range}") }
            }
            "insert_chart" => {
                let range = arg("range")?;
                let kind = arg("type")?;
                if es { format!("Insertar un gráfico ({kind}) de {range}") } else { format!("Insert a {kind} chart of {range}") }
            }
            "create_table" => {
                let range = arg("range")?;
                if es { format!("Crear una tabla con {range}") } else { format!("Make {range} a table") }
            }
            "sort_range" => {
                let range = arg("range").unwrap_or_else(|| if es { "la selección".into() } else { "the selection".into() });
                if es { format!("Ordenar {range}") } else { format!("Sort {range}") }
            }
            "filter" => {
                let column = arg("column")?;
                if es { format!("Filtrar la columna {column}") } else { format!("Filter column {column}") }
            }
            "run_command" => {
                let c = arg("command")?;
                let params = input.get("params").filter(|p| p.as_object().is_some_and(|o| !o.is_empty()));
                let params = params.map(|p| clip(&p.to_string(), PREVIEW)).unwrap_or_default();
                let label = self.app.session.commands().into_iter().find(|x| x.id == c).map(|x| x.label.to_string());
                let name = label.map_or_else(|| c.clone(), |l| format!("{l} ({c})"));
                match (es, params.is_empty()) {
                    (true, true) => format!("Ejecutar {name}"),
                    (true, false) => format!("Ejecutar {name} {params}"),
                    (false, true) => format!("Run {name}"),
                    (false, false) => format!("Run {name} {params}"),
                }
            }
            _ => return None,
        })
    }

    fn suggestions(&self, lang: Lang) -> Vec<String> {
        let s: [&str; 4] = if lang == Lang::Spanish {
            [
                "Resume lo que dicen estos datos",
                "Añade una columna de totales con fórmulas",
                "Crea un gráfico con estos datos",
                "Da formato de tabla profesional a la hoja",
            ]
        } else {
            ["Summarise what this data says", "Add a totals column with formulas", "Chart this data", "Format this sheet as a clean table"]
        };
        s.iter().map(|s| (*s).to_string()).collect()
    }

    fn undo_depth(&mut self) -> usize {
        self.app.session.active().map_or(0, |d| d.undo.len())
    }

    fn undo(&mut self, steps: usize) -> Result<(), String> {
        let before = self.undo_depth();
        self.run("edit.undo", json!({"steps": steps}))?;
        if before == self.undo_depth() && steps > 0 { Err("nothing to undo".into()) } else { Ok(()) }
    }
}

#[cfg(test)]
mod tests;
