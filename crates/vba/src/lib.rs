//! A VBA interpreter: the language and its own library, with the application's object model left to
//! a [`Host`].
//!
//! [`Project`] holds the modules of a workbook's `vbaProject.bin`; a [`Vm`] runs one of its
//! procedures against a host. Nothing here touches files, windows or the clock except through the
//! host, so the crate builds for the web.

#![forbid(unsafe_code)]
#![deny(clippy::unwrap_used, clippy::expect_used, clippy::panic, clippy::unimplemented, clippy::todo, clippy::unreachable)]

pub mod ast;
mod builtins;
pub mod debug;
mod format;
pub mod ide;
pub mod host;
mod interp;
mod lexer;
mod native;
mod ops;
mod parser;
pub mod project;
mod rewrite;
pub mod value;

pub use host::{CallArg, Host, NoHost, arg};
pub use interp::{DEFAULT_MAX_STEPS, ErrorSite, Vm};
pub use project::{ModuleKind, ProcInfo, Project, SourceModule};
pub use rewrite::rewrite_bin;
pub use value::{HostObject, Obj, VArray, Value, VbError};
