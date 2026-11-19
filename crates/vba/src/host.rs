//! What the interpreter needs from the program around it: the application's object model, dialogs
//! and files. A host implements [`Host`]; everything else about VBA lives in this crate.

use crate::ast::FileMode;
use crate::value::{Obj, Result, Value, VbError};

/// One argument of a call into the host: positional, or `Name:=value` (the name in lower case).
#[derive(Clone, Debug)]
pub struct CallArg {
    pub name: Option<String>,
    pub value: Value,
}

impl CallArg {
    pub fn positional(value: Value) -> CallArg {
        CallArg { name: None, value }
    }
}

/// The argument `index` of a call, whether given by position or by `name` (lower case).
pub fn arg<'a>(args: &'a [CallArg], index: usize, name: &str) -> Option<&'a Value> {
    if let Some(a) = args.iter().find(|a| a.name.as_deref() == Some(name)) {
        return Some(&a.value).filter(|v| !matches!(v, Value::Missing));
    }
    args.iter().filter(|a| a.name.is_none()).nth(index).map(|a| &a.value).filter(|v| !matches!(v, Value::Missing))
}

fn unsupported(what: &str) -> VbError {
    VbError::new(70, format!("{what} is not available here"))
}

pub trait Host {
    /// A property or method of `obj`. `name` is lower case. With `assign`, a property is being set
    /// (`obj.name(args) = value`); an empty `name` asks for the object's default member.
    fn member(&mut self, obj: &Obj, name: &str, args: &[CallArg], assign: Option<Value>) -> Result<Value>;

    /// A name that is not a variable or procedure of the macro: `Application`, `ActiveSheet`,
    /// `Range(...)`, `xlCenter`… `Ok(None)` means the host does not know it either.
    fn global(&mut self, name: &str, args: &[CallArg], assign: Option<Value>) -> Result<Option<Value>>;

    /// The items of an object for `For Each`.
    fn enumerate(&mut self, obj: &Obj) -> Result<Vec<Value>> {
        let _ = obj;
        Err(VbError::new(438, "This object cannot be enumerated"))
    }

    /// `CreateObject("Prog.Id")`.
    fn create_object(&mut self, prog_id: &str) -> Result<Value> {
        Err(VbError::new(429, format!("ActiveX component can't create object: '{prog_id}'")))
    }

    /// `MsgBox`: the button pressed (`vbOK` is 1).
    fn msg_box(&mut self, prompt: &str, buttons: i32, title: &str) -> i32 {
        let _ = (prompt, buttons, title);
        1
    }

    /// `InputBox`: the text typed, or `None` when it was cancelled.
    fn input_box(&mut self, prompt: &str, title: &str, default: &str) -> Option<String> {
        let _ = (prompt, title);
        Some(default.to_string())
    }

    fn file_open(&mut self, path: &str, mode: FileMode, number: i32) -> Result<()> {
        let _ = (path, mode, number);
        Err(unsupported("File access"))
    }

    /// Appends `text` to the file open as `number`.
    fn file_write(&mut self, number: i32, text: &str) -> Result<()> {
        let _ = (number, text);
        Err(unsupported("File access"))
    }

    /// The next line of the file, or `None` at its end.
    fn file_read_line(&mut self, number: i32) -> Result<Option<String>> {
        let _ = number;
        Err(unsupported("File access"))
    }

    /// `EOF(number)`: whether the file open for input has nothing left to read.
    fn file_eof(&mut self, number: i32) -> Result<bool> {
        let _ = number;
        Err(unsupported("File access"))
    }

    /// Closes file `number`, or every open file with `None`.
    fn file_close(&mut self, number: Option<i32>) -> Result<()> {
        let _ = number;
        Ok(())
    }

    /// `Dir(pattern, attributes)`; an empty pattern continues the previous search.
    fn dir(&mut self, pattern: &str, attributes: i32) -> Result<String> {
        let _ = (pattern, attributes);
        Err(unsupported("File access"))
    }

    fn make_dir(&mut self, path: &str) -> Result<()> {
        let _ = path;
        Err(unsupported("File access"))
    }

    fn remove_file(&mut self, path: &str) -> Result<()> {
        let _ = path;
        Err(unsupported("File access"))
    }

    /// `Environ(name)`.
    fn environ(&mut self, name: &str) -> String {
        let _ = name;
        String::new()
    }

    /// Called now and then while a macro runs; an error stops it (a Cancel button, a time limit).
    fn tick(&mut self) -> Result<()> {
        Ok(())
    }

    /// The current time as a serial number (the 1900 date system).
    fn now(&mut self) -> f64 {
        celdas_core::date::unix_now_ms().map_or(0.0, celdas_core::date::serial_from_unix_ms)
    }
}

/// A host that knows nothing: for running pure VBA code (and the tests of it).
pub struct NoHost;

impl Host for NoHost {
    fn member(&mut self, obj: &Obj, name: &str, _args: &[CallArg], _assign: Option<Value>) -> Result<Value> {
        let _ = obj;
        Err(VbError::no_member(name))
    }
    fn global(&mut self, _name: &str, _args: &[CallArg], _assign: Option<Value>) -> Result<Option<Value>> {
        Ok(None)
    }
}
