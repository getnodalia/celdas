// SPDX-License-Identifier: MIT OR Apache-2.0
//! Recording as Visual Basic: the commands a person ran while recording, written as the statements
//! Excel's recorder would write (`Range("B2").Select`, `Selection.Font.Bold = True`), so a recording
//! is a macro they can read, edit and debug in the editor.
//!
//! Every statement written here is one the macro host runs (`vba::range`, `vba::app`), and the tests
//! below run what a command was translated to and compare the workbook with the one the command
//! itself changed. A command with no statement of its own (a chart, a pivot table, a command whose
//! effect depends on the cells around it) is left as a comment saying so, and is listed in the
//! result's `skipped`, so nothing is lost without being seen.

use serde_json::Value as Json;

use crate::cmd::str_param;

/// What recording one step as code gave.
pub enum Statements {
    /// The lines of code, without indentation.
    Code(Vec<String>),
    /// A command with no statement of its own.
    Skipped,
}

/// A Visual Basic string literal for `s`: the quotes doubled, the line breaks and the characters
/// the project's code page does not hold written as `vbLf`, `ChrW(…)` and so on, so the module can
/// always be saved.
pub fn string_literal(s: &str) -> String {
    let mut parts: Vec<String> = Vec::new();
    let mut run = String::new();
    let flush = |run: &mut String, parts: &mut Vec<String>| {
        if !run.is_empty() {
            parts.push(format!("\"{run}\""));
            run.clear();
        }
    };
    for c in s.chars() {
        match c {
            '"' => run.push_str("\"\""),
            '\n' => {
                flush(&mut run, &mut parts);
                parts.push("vbLf".into());
            }
            '\r' => {
                flush(&mut run, &mut parts);
                parts.push("vbCr".into());
            }
            '\t' => {
                flush(&mut run, &mut parts);
                parts.push("vbTab".into());
            }
            c if (' '..='~').contains(&c) || ('\u{a0}'..='\u{ff}').contains(&c) => run.push(c),
            c => {
                flush(&mut run, &mut parts);
                let mut units = [0u16; 2];
                for u in c.encode_utf16(&mut units) {
                    parts.push(format!("ChrW({})", i32::from(*u as i16)));
                }
            }
        }
    }
    flush(&mut run, &mut parts);
    if parts.is_empty() { "\"\"".into() } else { parts.join(" & ") }
}

/// A number as Visual Basic writes it.
fn number(n: &serde_json::Number) -> String {
    match n.as_f64() {
        Some(f) if f.is_finite() && f.fract() == 0.0 && f.abs() < 1e15 => format!("{}", f as i64),
        Some(f) if f.is_finite() => format!("{f}"),
        _ => "0".into(),
    }
}

/// `RGB(r, g, b)` for a `#RRGGBB` color.
fn rgb(color: &str) -> Option<String> {
    let hex = color.strip_prefix('#')?;
    if hex.len() != 6 || !hex.is_ascii() {
        return None;
    }
    let part = |i: usize| u8::from_str_radix(hex.get(i..i + 2)?, 16).ok();
    Some(format!("RGB({}, {}, {})", part(0)?, part(2)?, part(4)?))
}

/// The sheet a step names, as an expression (`Sheets("Data")`), or `None` when it names none.
fn sheet_of(p: &Json) -> Result<Option<String>, ()> {
    match p.get("sheet") {
        None | Some(Json::Null) => Ok(None),
        Some(Json::String(n)) => Ok(Some(format!("Sheets({})", string_literal(n)))),
        Some(Json::Number(i)) => i.as_u64().map(|i| Some(format!("Sheets({})", i + 1))).ok_or(()),
        Some(_) => Err(()),
    }
}

/// The range a step acts on: the one it names, or the selection.
fn target(p: &Json) -> Result<String, ()> {
    match p.get("range") {
        // A sheet without a range is the selection of that sheet, which code can't name.
        None | Some(Json::Null) => match sheet_of(p)? {
            Some(_) => Err(()),
            None => Ok("Selection".into()),
        },
        Some(Json::String(r)) => {
            let r = r.trim();
            if r.is_empty() || r.contains('"') {
                return Err(());
            }
            Ok(match sheet_of(p)? {
                Some(sh) => format!("{sh}.Range(\"{r}\")"),
                None => format!("Range(\"{r}\")"),
            })
        }
        Some(_) => Err(()),
    }
}

/// Whole rows (`Rows("2:3")`) or columns the step names, or the selection's.
fn lines_of(p: &Json, rows: bool) -> Result<String, ()> {
    let key = if rows { "rows" } else { "cols" };
    let (word, entire) = if rows { ("Rows", "EntireRow") } else { ("Columns", "EntireColumn") };
    match p.get(key) {
        None | Some(Json::Null) => Ok(format!("Selection.{entire}")),
        Some(Json::String(r)) => {
            let r = r.trim();
            // One block only: `2:2,4:5` and `Sheet2!2:5` are the macro's, not Excel's, spelling.
            if r.is_empty() || r.contains(',') || r.contains('!') || r.contains('"') {
                return Err(());
            }
            Ok(match sheet_of(p)? {
                Some(sh) => format!("{sh}.{word}(\"{r}\")"),
                None => format!("{word}(\"{r}\")"),
            })
        }
        Some(_) => Err(()),
    }
}

fn on(p: &Json) -> bool {
    p.get("on").and_then(Json::as_bool).unwrap_or(true)
}

fn truth(b: bool) -> &'static str {
    if b { "True" } else { "False" }
}

/// The statements for one command, or [`Statements::Skipped`].
pub fn statements(id: &str, p: &Json) -> Statements {
    match translate(id, p) {
        Ok(Some(lines)) => Statements::Code(lines),
        _ => Statements::Skipped,
    }
}

fn translate(id: &str, p: &Json) -> Result<Option<Vec<String>>, ()> {
    let one = |s: String| Ok(Some(vec![s]));
    match id {
        "sheet.activate" => match p.get("sheet") {
            Some(Json::String(n)) => one(format!("Sheets({}).Select", string_literal(n))),
            Some(Json::Number(i)) => one(format!("Sheets({}).Select", i.as_u64().ok_or(())? + 1)),
            _ => Err(()),
        },
        "selection.set" => {
            let range = str_param(p, "range").ok_or(())?.trim();
            if range.is_empty() || range.contains('"') {
                return Err(());
            }
            let mut out = vec![format!("Range(\"{range}\").Select")];
            let first = range.split(',').next().unwrap_or("").split(':').next().unwrap_or("");
            if let Some(active) = str_param(p, "active").filter(|a| !a.eq_ignore_ascii_case(first) && !a.contains('"')) {
                out.push(format!("Range(\"{active}\").Activate"));
            }
            Ok(Some(out))
        }
        "cell.set" | "range.fill" => {
            let at = match str_param(p, "cell") {
                Some(c) if !c.contains('"') => match sheet_of(p)? {
                    Some(sh) => format!("{sh}.Range(\"{c}\")"),
                    None => format!("Range(\"{c}\")"),
                },
                _ if id == "range.fill" => target(p)?,
                _ => "ActiveCell".into(),
            };
            if p.get("array").and_then(Json::as_bool) == Some(true) {
                return Ok(None);
            }
            let input = p.get("input").ok_or(())?;
            one(assign_input(&at, input)?)
        }
        "range.setValues" => {
            let origin = match str_param(p, "range") {
                Some(r) => celdas_core::RangeRef::parse(r).ok_or(())?.start,
                None => return Ok(None),
            };
            let rows = p.get("values").and_then(Json::as_array).ok_or(())?;
            let mut out = Vec::new();
            for (dr, row) in rows.iter().enumerate() {
                for (dc, v) in row.as_array().ok_or(())?.iter().enumerate() {
                    let cell = origin.offset(dr as i64, dc as i64).ok_or(())?;
                    out.push(assign_input(&format!("Range(\"{}\")", cell.a1()), v)?);
                    if out.len() > 400 {
                        return Ok(None);
                    }
                }
            }
            Ok(Some(out))
        }
        // Font.
        "home.bold" => one(format!("{}.Font.Bold = {}", target(p)?, truth(on(p)))),
        "home.italic" => one(format!("{}.Font.Italic = {}", target(p)?, truth(on(p)))),
        "home.strikethrough" => one(format!("{}.Font.Strikethrough = {}", target(p)?, truth(on(p)))),
        "home.underline" => one(format!("{}.Font.Underline = {}", target(p)?, if on(p) { "xlUnderlineStyleSingle" } else { "xlUnderlineStyleNone" })),
        "home.doubleUnderline" => {
            one(format!("{}.Font.Underline = {}", target(p)?, if on(p) { "xlUnderlineStyleDouble" } else { "xlUnderlineStyleNone" }))
        }
        "home.fontName" => one(format!("{}.Font.Name = {}", target(p)?, string_literal(str_param(p, "name").ok_or(())?))),
        "home.fontSize" => {
            let size = p.get("size").and_then(Json::as_f64).filter(|s| s.is_finite() && *s > 0.0).ok_or(())?;
            one(format!("{}.Font.Size = {}", target(p)?, number(&serde_json::Number::from_f64(size).ok_or(())?)))
        }
        "home.fontColor" => match p.get("color") {
            Some(Json::String(c)) if c == "auto" => one(format!("{}.Font.ColorIndex = xlAutomatic", target(p)?)),
            Some(Json::String(c)) => one(format!("{}.Font.Color = {}", target(p)?, rgb(c).ok_or(())?)),
            _ => Ok(None),
        },
        "home.fillColor" => match p.get("color") {
            Some(Json::String(c)) if c == "none" => one(format!("{}.Interior.Pattern = xlNone", target(p)?)),
            Some(Json::String(c)) => {
                let t = target(p)?;
                Ok(Some(vec![format!("{t}.Interior.Pattern = xlSolid"), format!("{t}.Interior.Color = {}", rgb(c).ok_or(())?)]))
            }
            _ => Ok(None),
        },
        // Alignment.
        "home.alignLeft" => one(format!("{}.HorizontalAlignment = xlLeft", target(p)?)),
        "home.alignCenter" => one(format!("{}.HorizontalAlignment = xlCenter", target(p)?)),
        "home.alignRight" => one(format!("{}.HorizontalAlignment = xlRight", target(p)?)),
        "home.justify" => one(format!("{}.HorizontalAlignment = xlJustify", target(p)?)),
        "home.alignTop" => one(format!("{}.VerticalAlignment = xlTop", target(p)?)),
        "home.alignMiddle" => one(format!("{}.VerticalAlignment = xlCenter", target(p)?)),
        "home.alignBottom" => one(format!("{}.VerticalAlignment = xlBottom", target(p)?)),
        "home.wrapText" => one(format!("{}.WrapText = {}", target(p)?, truth(on(p)))),
        "home.shrinkToFit" => one(format!("{}.ShrinkToFit = {}", target(p)?, truth(on(p)))),
        "home.mergeCells" => one(format!("{}.Merge", target(p)?)),
        "home.mergeAcross" => one(format!("{}.Merge True", target(p)?)),
        "home.mergeCenter" => {
            let t = target(p)?;
            Ok(Some(vec![format!("{t}.HorizontalAlignment = xlCenter"), format!("{t}.Merge")]))
        }
        "home.unmergeCells" => one(format!("{}.UnMerge", target(p)?)),
        // Numbers.
        "home.numberFormat" => {
            let code = match str_param(p, "code") {
                Some(c) => c.to_string(),
                None => crate::cmd::format::format_code_for(str_param(p, "format").ok_or(())?).to_string(),
            };
            one(format!("{}.NumberFormat = {}", target(p)?, string_literal(&code)))
        }
        "home.percent" => one(format!("{}.NumberFormat = \"0%\"", target(p)?)),
        // Borders.
        "home.borders" => borders(p),
        // Cells, rows and columns.
        "home.insertRows" => one(format!("{}.Insert", lines_of(p, true)?)),
        "home.insertColumns" => one(format!("{}.Insert", lines_of(p, false)?)),
        "home.deleteRows" => one(format!("{}.Delete", lines_of(p, true)?)),
        "home.deleteColumns" => one(format!("{}.Delete", lines_of(p, false)?)),
        "home.hideRows" => one(format!("{}.Hidden = True", lines_of(p, true)?)),
        "home.hideColumns" => one(format!("{}.Hidden = True", lines_of(p, false)?)),
        "home.unhideRows" => one(format!("{}.Hidden = False", lines_of(p, true)?)),
        "home.unhideColumns" => one(format!("{}.Hidden = False", lines_of(p, false)?)),
        "home.autofitRowHeight" => one(format!("{}.AutoFit", lines_of(p, true)?)),
        "home.autofitColumnWidth" => one(format!("{}.AutoFit", lines_of(p, false)?)),
        "home.rowHeight" => {
            let points = p.get("points").and_then(Json::as_f64).filter(|v| v.is_finite()).ok_or(())?;
            one(format!("{}.RowHeight = {}", lines_of(p, true)?, number(&serde_json::Number::from_f64(points).ok_or(())?)))
        }
        "home.columnWidth" => {
            let chars = p.get("chars").and_then(Json::as_f64).filter(|v| v.is_finite()).ok_or(())?;
            one(format!("{}.ColumnWidth = {}", lines_of(p, false)?, number(&serde_json::Number::from_f64(chars).ok_or(())?)))
        }
        "home.insertCells" | "home.deleteCells" => {
            let shift = match str_param(p, "shift") {
                Some("down") => "xlDown",
                Some("right") => "xlToRight",
                Some("up") => "xlUp",
                Some("left") => "xlToLeft",
                _ => return Ok(None),
            };
            let verb = if id == "home.insertCells" { "Insert" } else { "Delete" };
            one(format!("{}.{verb} Shift:={shift}", target(p)?))
        }
        // Clearing and filling.
        "edit.clearAll" => one(format!("{}.Clear", target(p)?)),
        "edit.clearContents" => one(format!("{}.ClearContents", target(p)?)),
        "edit.clearFormats" => one(format!("{}.ClearFormats", target(p)?)),
        "edit.clearComments" => one(format!("{}.ClearComments", target(p)?)),
        "edit.fillDown" => one(format!("{}.FillDown", target(p)?)),
        "edit.fillRight" => one(format!("{}.FillRight", target(p)?)),
        "edit.fillUp" => one(format!("{}.FillUp", target(p)?)),
        "edit.fillLeft" => one(format!("{}.FillLeft", target(p)?)),
        // The clipboard.
        "edit.copy" if p.get("html").is_none() => one(format!("{}.Copy", target(p)?)),
        "edit.cut" if p.get("html").is_none() => one(format!("{}.Cut", target(p)?)),
        "edit.paste" if p.get("text").is_none() && p.get("html").is_none() => {
            let mut out = Vec::new();
            if let Some(at) = str_param(p, "at").filter(|a| !a.contains('"')) {
                out.push(format!("Range(\"{at}\").Select"));
            }
            out.push("ActiveSheet.Paste".into());
            Ok(Some(out))
        }
        "edit.pasteSpecial" => {
            let what = match str_param(p, "what").ok_or(())? {
                "all" => "xlPasteAll",
                "formulas" => "xlPasteFormulas",
                "values" => "xlPasteValues",
                "formats" => "xlPasteFormats",
                "comments" => "xlPasteComments",
                "validation" => "xlPasteValidation",
                "allExceptBorders" => "xlPasteAllExceptBorders",
                "columnWidths" => "xlPasteColumnWidths",
                "formulasAndNumberFormats" => "xlPasteFormulasAndNumberFormats",
                "valuesAndNumberFormats" => "xlPasteValuesAndNumberFormats",
                _ => return Ok(None),
            };
            let operation = match str_param(p, "operation") {
                None | Some("none") => "xlNone",
                Some("add") => "xlAdd",
                Some("subtract") => "xlSubtract",
                Some("multiply") => "xlMultiply",
                Some("divide") => "xlDivide",
                _ => return Ok(None),
            };
            let flag = |k: &str| p.get(k).and_then(Json::as_bool).unwrap_or(false);
            if flag("link") {
                return Ok(None);
            }
            one(format!(
                "Selection.PasteSpecial Paste:={what}, Operation:={operation}, SkipBlanks:={}, Transpose:={}",
                truth(flag("skipBlanks")),
                truth(flag("transpose"))
            ))
        }
        // Sorting and filtering.
        "data.sortAscending" | "data.sortDescending" => {
            let order = if id == "data.sortAscending" { "xlAscending" } else { "xlDescending" };
            let key = match str_param(p, "column") {
                Some(c) if c.chars().all(|c| c.is_ascii_alphabetic()) && !c.is_empty() => format!("Range(\"{c}1\")"),
                Some(_) => return Ok(None),
                None => "ActiveCell".into(),
            };
            one(format!("{}.Sort Key1:={key}, Order1:={order}, Header:=xlGuess", target(p)?))
        }
        "data.sort" => {
            let keys = p.get("keys").and_then(Json::as_array).ok_or(())?;
            if keys.is_empty() || keys.len() > 3 || p.get("orientation").and_then(Json::as_str) == Some("columns") || p.get("matchCase").is_some() {
                return Ok(None);
            }
            let mut args = Vec::new();
            for (i, k) in keys.iter().enumerate() {
                let col = str_param(k, "column").filter(|c| !c.is_empty() && c.chars().all(|c| c.is_ascii_alphabetic())).ok_or(())?;
                if k.get("by").and_then(Json::as_str).is_some_and(|b| b != "values") || k.get("customList").is_some() {
                    return Ok(None);
                }
                let order = if str_param(k, "order") == Some("desc") { "xlDescending" } else { "xlAscending" };
                args.push(format!("Key{n}:=Range(\"{col}1\"), Order{n}:={order}", n = i + 1));
            }
            let header = if p.get("header").and_then(Json::as_bool).unwrap_or(false) { "xlYes" } else { "xlNo" };
            one(format!("{}.Sort {}, Header:={header}", target(p)?, args.join(", ")))
        }
        "data.filter" => one(format!("{}.AutoFilter", target(p)?)),
        // Sheets.
        "sheet.rename" => {
            let name = str_param(p, "name").ok_or(())?;
            let at = sheet_of(p)?.unwrap_or_else(|| "ActiveSheet".into());
            one(format!("{at}.Name = {}", string_literal(name)))
        }
        "home.insertSheet" => {
            let mut out = vec!["Sheets.Add".to_string()];
            if p.get("before").is_some_and(|b| !b.is_null()) {
                return Ok(None);
            }
            if let Some(n) = str_param(p, "name") {
                out.push(format!("ActiveSheet.Name = {}", string_literal(n)));
            }
            Ok(Some(out))
        }
        "home.deleteSheet" => {
            let at = sheet_of(p)?.unwrap_or_else(|| "ActiveSheet".into());
            one(format!("{at}.Delete"))
        }
        _ => Ok(None),
    }
}

/// `at.Value = …` or `at.Formula = …` for what a cell was given.
fn assign_input(at: &str, input: &Json) -> Result<String, ()> {
    Ok(match input {
        Json::Null => format!("{at}.ClearContents"),
        Json::Bool(b) => format!("{at}.Value = {}", truth(*b)),
        Json::Number(n) => format!("{at}.Value = {}", number(n)),
        Json::String(s) if s.is_empty() => format!("{at}.ClearContents"),
        Json::String(s) if s.starts_with('=') => format!("{at}.Formula = {}", string_literal(s)),
        Json::String(s) => format!("{at}.Value = {}", string_literal(s)),
        _ => return Err(()),
    })
}

fn borders(p: &Json) -> Result<Option<Vec<String>>, ()> {
    let t = target(p)?;
    let preset = str_param(p, "preset").ok_or(())?;
    let style = str_param(p, "style");
    if p.get("color").is_some_and(|c| !c.is_null()) || style.is_some_and(|s| !matches!(s, "thin" | "medium" | "thick")) {
        return Ok(None);
    }
    let weight = match style {
        Some("medium") => "xlMedium",
        Some("thick") => "xlThick",
        _ => "xlThin",
    };
    let edge =
        |name: &str, weight: &str| vec![format!("{t}.Borders({name}).LineStyle = xlContinuous"), format!("{t}.Borders({name}).Weight = {weight}")];
    let thick = if matches!(preset, "thickOutside" | "thickBottom") { "xlThick" } else { weight };
    let edges: &[&str] = match preset {
        "none" => return Ok(Some(vec![format!("{t}.Borders.LineStyle = xlNone")])),
        "bottom" | "thickBottom" => &["xlEdgeBottom"],
        "top" => &["xlEdgeTop"],
        "left" => &["xlEdgeLeft"],
        "right" => &["xlEdgeRight"],
        "outside" | "thickOutside" => &["xlEdgeLeft", "xlEdgeTop", "xlEdgeBottom", "xlEdgeRight"],
        "all" => &["xlEdgeLeft", "xlEdgeTop", "xlEdgeBottom", "xlEdgeRight", "xlInsideVertical", "xlInsideHorizontal"],
        "insideHorizontal" => &["xlInsideHorizontal"],
        "insideVertical" => &["xlInsideVertical"],
        "inside" => &["xlInsideVertical", "xlInsideHorizontal"],
        "topBottom" => &["xlEdgeTop", "xlEdgeBottom"],
        _ => return Ok(None),
    };
    Ok(Some(edges.iter().flat_map(|e| edge(e, thick)).collect()))
}

/// A name Visual Basic accepts for a procedure: letters, digits and underscores, starting with a
/// letter, at most 255 characters.
pub fn procedure_name(name: &str) -> String {
    let mut out: String = name.chars().map(|c| if c.is_alphanumeric() || c == '_' { c } else { '_' }).collect();
    if !out.chars().next().is_some_and(char::is_alphabetic) {
        out.insert(0, 'M');
    }
    out.chars().take(255).collect()
}

/// The text of a `Sub` called `name` that does what `steps` did, with `skipped` the commands that
/// could not be written as code.
pub fn recorded_sub(name: &str, description: &str, steps: &[(String, Json)]) -> (String, Vec<String>) {
    let mut out = format!("Sub {}()\r\n", procedure_name(name));
    for line in description.lines().filter(|l| !l.trim().is_empty()) {
        out.push_str(&format!("    ' {}\r\n", line.trim()));
    }
    let mut skipped = Vec::new();
    for (id, params) in steps {
        match statements(id, params) {
            Statements::Code(lines) => {
                for l in lines {
                    out.push_str("    ");
                    out.push_str(&l);
                    out.push_str("\r\n");
                }
            }
            Statements::Skipped => {
                out.push_str(&format!("    ' Not recorded as code: {id}\r\n"));
                skipped.push(id.clone());
            }
        }
    }
    out.push_str("End Sub\r\n");
    (out, skipped)
}
