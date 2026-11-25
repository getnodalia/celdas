// SPDX-License-Identifier: MIT OR Apache-2.0
//! What an editor needs to know about VBA source: the structure of a module, the mistakes in it, the
//! names it can offer, and how to lay it out. Everything here reads source the way a person sees it
//! while typing, so it works on modules that do not compile.

pub mod analyze;
pub mod catalog;
pub mod complete;
pub mod format;
pub mod lines;
pub mod scan;

pub use analyze::{Analysis, analyze};
pub use catalog::{Catalog, Class, Global, GlobalKind, Member, MemberKind};
pub use format::format;
pub use scan::{Diagnostic, ParamDecl, ProcDecl, Scan, Severity, TypeDecl, VarDecl, scan};
