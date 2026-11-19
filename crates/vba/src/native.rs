//! Objects VBA itself provides: `Err`, `Debug`, `Collection` and `Scripting.Dictionary`.

use std::any::Any;
use std::cell::RefCell;

use crate::host::{arg, CallArg};
use crate::value::{HostObject, Result, Value, VbError};

pub struct ErrObject;

impl HostObject for ErrObject {
    fn class(&self) -> String {
        "ErrObject".into()
    }
    fn as_any(&self) -> &dyn Any {
        self
    }
    fn same(&self, other: &dyn HostObject) -> bool {
        other.as_any().is::<ErrObject>()
    }
}

pub struct DebugObject;

impl HostObject for DebugObject {
    fn class(&self) -> String {
        "Debug".into()
    }
    fn as_any(&self) -> &dyn Any {
        self
    }
    fn same(&self, other: &dyn HostObject) -> bool {
        other.as_any().is::<DebugObject>()
    }
}

/// A VBA `Collection`: items in order, optionally with a key.
#[derive(Default)]
pub struct Collection {
    items: RefCell<Vec<(Option<String>, Value)>>,
}

impl Collection {
    pub fn items(&self) -> Vec<Value> {
        self.items.borrow().iter().map(|(_, v)| v.clone()).collect()
    }

    fn position(&self, key: &Value) -> Result<usize> {
        let items = self.items.borrow();
        match key {
            Value::Str(s) => items.iter().position(|(k, _)| k.as_deref().is_some_and(|k| k.eq_ignore_ascii_case(s))).ok_or_else(|| VbError::new(5, "Invalid procedure call or argument")),
            other => {
                let i = other.to_i32()?;
                if i >= 1 && (i as usize) <= items.len() { Ok(i as usize - 1) } else { Err(VbError::subscript()) }
            }
        }
    }

    pub fn member(&self, name: &str, args: &[CallArg]) -> Result<Value> {
        match name {
            "add" => {
                let item = arg(args, 0, "item").cloned().unwrap_or(Value::Empty);
                let key = match arg(args, 1, "key") {
                    Some(k) => Some(k.to_text()?),
                    None => None,
                };
                if let Some(k) = &key
                    && self.items.borrow().iter().any(|(e, _)| e.as_deref().is_some_and(|e| e.eq_ignore_ascii_case(k)))
                {
                    return Err(VbError::new(457, "This key is already associated with an element of this collection"));
                }
                let mut items = self.items.borrow_mut();
                // Before / After
                if let Some(before) = arg(args, 2, "before") {
                    let at = match before {
                        Value::Str(_) => {
                            drop(items);
                            let p = self.position(before)?;
                            items = self.items.borrow_mut();
                            p
                        }
                        n => (n.to_i32()?.max(1) as usize).saturating_sub(1).min(items.len()),
                    };
                    items.insert(at, (key, item));
                } else {
                    items.push((key, item));
                }
                Ok(Value::Empty)
            }
            "count" => Ok(Value::Long(self.items.borrow().len() as i32)),
            "" | "item" => {
                let key = arg(args, 0, "index").ok_or_else(VbError::invalid_call)?;
                let p = self.position(key)?;
                Ok(self.items.borrow().get(p).map(|(_, v)| v.clone()).unwrap_or(Value::Empty))
            }
            "remove" => {
                let key = arg(args, 0, "index").ok_or_else(VbError::invalid_call)?;
                let p = self.position(key)?;
                self.items.borrow_mut().remove(p);
                Ok(Value::Empty)
            }
            other => Err(VbError::no_member(other)),
        }
    }
}

impl HostObject for Collection {
    fn class(&self) -> String {
        "Collection".into()
    }
    fn as_any(&self) -> &dyn Any {
        self
    }
    fn same(&self, _other: &dyn HostObject) -> bool {
        false
    }
}

/// A `Scripting.Dictionary`: keys keep their insertion order; text keys ignore case unless asked.
#[derive(Default)]
pub struct Dictionary {
    entries: RefCell<Vec<(Value, Value)>>,
    text_compare: RefCell<bool>,
}

impl Dictionary {
    pub fn keys(&self) -> Vec<Value> {
        self.entries.borrow().iter().map(|(k, _)| k.clone()).collect()
    }

    fn find(&self, key: &Value) -> Option<usize> {
        let text = *self.text_compare.borrow();
        self.entries.borrow().iter().position(|(k, _)| match (k, key) {
            (Value::Str(a), Value::Str(b)) => if text { a.eq_ignore_ascii_case(b) } else { a == b },
            (a, b) => a == b,
        })
    }

    pub fn member(&self, name: &str, args: &[CallArg], assign: Option<Value>) -> Result<Value> {
        match name {
            "add" => {
                let key = arg(args, 0, "key").cloned().ok_or_else(VbError::invalid_call)?;
                let item = arg(args, 1, "item").cloned().unwrap_or(Value::Empty);
                if self.find(&key).is_some() {
                    return Err(VbError::new(457, "This key is already associated with an element of this collection"));
                }
                self.entries.borrow_mut().push((key, item));
                Ok(Value::Empty)
            }
            "exists" => {
                let key = arg(args, 0, "key").ok_or_else(VbError::invalid_call)?;
                Ok(Value::Bool(self.find(key).is_some()))
            }
            "" | "item" => {
                let key = arg(args, 0, "key").cloned().ok_or_else(VbError::invalid_call)?;
                if let Some(v) = assign {
                    match self.find(&key) {
                        Some(i) => {
                            if let Some(slot) = self.entries.borrow_mut().get_mut(i) {
                                slot.1 = v;
                            }
                        }
                        None => self.entries.borrow_mut().push((key, v)),
                    }
                    return Ok(Value::Empty);
                }
                // Reading a missing key adds it, empty, as the real object does.
                if let Some(i) = self.find(&key) {
                    return Ok(self.entries.borrow().get(i).map(|(_, v)| v.clone()).unwrap_or(Value::Empty));
                }
                self.entries.borrow_mut().push((key, Value::Empty));
                Ok(Value::Empty)
            }
            "count" => Ok(Value::Long(self.entries.borrow().len() as i32)),
            "remove" => {
                let key = arg(args, 0, "key").ok_or_else(VbError::invalid_call)?;
                let i = self.find(key).ok_or_else(|| VbError::new(32811, "Application-defined or object-defined error"))?;
                self.entries.borrow_mut().remove(i);
                Ok(Value::Empty)
            }
            "removeall" => {
                self.entries.borrow_mut().clear();
                Ok(Value::Empty)
            }
            "keys" => Ok(Value::array(crate::value::VArray::from_vec(self.keys(), 0))),
            "items" => Ok(Value::array(crate::value::VArray::from_vec(self.entries.borrow().iter().map(|(_, v)| v.clone()).collect(), 0))),
            "comparemode" => {
                if let Some(v) = assign {
                    *self.text_compare.borrow_mut() = v.to_i32()? == 1;
                    return Ok(Value::Empty);
                }
                Ok(Value::Long(0))
            }
            other => Err(VbError::no_member(other)),
        }
    }
}

impl HostObject for Dictionary {
    fn class(&self) -> String {
        "Dictionary".into()
    }
    fn as_any(&self) -> &dyn Any {
        self
    }
    fn same(&self, _other: &dyn HostObject) -> bool {
        false
    }
}
