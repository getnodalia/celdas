//! The syntax tree of a VBA module. Names are lower-cased: VBA is not case-sensitive.

use std::rc::Rc;

use crate::value::Value;

pub type Name = Rc<str>;

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum Un {
    Neg,
    Plus,
    Not,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum Bin {
    Add,
    Sub,
    Mul,
    Div,
    IntDiv,
    Mod,
    Pow,
    Concat,
    Eq,
    Ne,
    Lt,
    Gt,
    Le,
    Ge,
    Like,
    Is,
    And,
    Or,
    Xor,
    Eqv,
    Imp,
}

/// One argument of a call: positional or `Name:=value`, or omitted (`f(a, , c)`).
#[derive(Clone, Debug)]
pub struct Arg {
    pub name: Option<Name>,
    pub value: Option<Expr>,
}

#[derive(Clone, Debug)]
pub enum Expr {
    Lit(Value),
    /// A variable, a constant, or a procedure called without arguments.
    Name(Name),
    /// `.name` inside a `With` block.
    WithMember(Name),
    Member(Box<Expr>, Name),
    /// `target(args)`: an array element, a function call, or a parameterized member.
    Call(Box<Expr>, Vec<Arg>),
    Unary(Un, Box<Expr>),
    Binary(Bin, Box<Expr>, Box<Expr>),
    New(Name),
    TypeOf(Box<Expr>, Name),
}

#[derive(Clone, Debug, PartialEq)]
pub enum VType {
    Variant,
    Boolean,
    Byte,
    Integer,
    Long,
    Single,
    Double,
    Currency,
    Date,
    String,
    /// Any class: `Worksheet`, `Range`, `Object`, `Collection`…
    Object(Name),
}

impl VType {
    pub fn from_name(name: &str) -> VType {
        match name {
            "variant" => VType::Variant,
            "boolean" => VType::Boolean,
            "byte" => VType::Byte,
            "integer" => VType::Integer,
            "long" | "longlong" | "longptr" => VType::Long,
            "single" => VType::Single,
            "double" | "decimal" => VType::Double,
            "currency" => VType::Currency,
            "date" => VType::Date,
            "string" => VType::String,
            other => VType::Object(Rc::from(other)),
        }
    }
}

#[derive(Clone, Debug)]
pub struct DimItem {
    pub name: Name,
    pub ty: VType,
    /// `Dim a(1 To 5)`: lower and upper bound per dimension (the lower is `None` when omitted).
    pub bounds: Option<Vec<(Option<Expr>, Expr)>>,
    /// `Dim a()`: a dynamic array.
    pub dynamic: bool,
    /// `As New Class`.
    pub is_new: bool,
}

#[derive(Clone, Debug)]
pub enum CaseTest {
    Value(Expr),
    Range(Expr, Expr),
    Is(Bin, Expr),
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum ExitKind {
    Sub,
    Function,
    Property,
    For,
    Do,
}

#[derive(Clone, Debug)]
pub enum OnError {
    ResumeNext,
    Goto0,
    Goto(Name),
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum FileMode {
    Input,
    Output,
    Append,
    Binary,
    Random,
}

#[derive(Clone, Debug)]
pub struct Stmt {
    pub line: u32,
    pub kind: StmtKind,
}

#[derive(Clone, Debug)]
pub enum StmtKind {
    Dim {
        items: Vec<DimItem>,
        is_static: bool,
    },
    Const(Vec<(Name, Expr)>),
    ReDim {
        preserve: bool,
        items: Vec<DimItem>,
    },
    Assign(Expr, Expr),
    SetAssign(Expr, Expr),
    Call(Expr),
    If {
        branches: Vec<(Expr, Vec<Stmt>)>,
        otherwise: Option<Vec<Stmt>>,
    },
    For {
        var: Expr,
        from: Expr,
        to: Expr,
        step: Option<Expr>,
        body: Vec<Stmt>,
    },
    ForEach {
        var: Expr,
        collection: Expr,
        body: Vec<Stmt>,
    },
    /// `Do [While|Until c] … Loop [While|Until c]`: `until` flips the sense of `cond`; `post`
    /// says the test comes after the body.
    Do {
        cond: Option<Expr>,
        until: bool,
        post: bool,
        body: Vec<Stmt>,
    },
    While(Expr, Vec<Stmt>),
    Select {
        subject: Expr,
        cases: Vec<(Vec<CaseTest>, Vec<Stmt>)>,
        otherwise: Option<Vec<Stmt>>,
    },
    With(Expr, Vec<Stmt>),
    Exit(ExitKind),
    OnError(OnError),
    Goto(Name),
    Label(Name),
    Resume(bool),
    Open {
        path: Expr,
        mode: FileMode,
        number: Expr,
    },
    Close(Vec<Expr>),
    Print {
        number: Expr,
        items: Vec<Expr>,
        newline: bool,
        quoted: bool,
    },
    LineInput {
        number: Expr,
        target: Expr,
    },
    Erase(Vec<Expr>),
    End,
    Nop,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum ProcKind {
    Sub,
    Function,
    PropertyGet,
    PropertyLet,
    PropertySet,
}

#[derive(Clone, Debug)]
pub struct Param {
    pub name: Name,
    pub by_val: bool,
    pub optional: bool,
    pub param_array: bool,
    pub ty: VType,
    pub default: Option<Expr>,
}

#[derive(Clone, Debug)]
pub struct Proc {
    pub name: Name,
    /// The name as written, for messages and lists.
    pub display: String,
    pub kind: ProcKind,
    pub public: bool,
    pub params: Vec<Param>,
    pub returns: VType,
    pub body: Vec<Stmt>,
    pub line: u32,
}

/// A module-level `Dim`/`Public`/`Const`.
#[derive(Clone, Debug)]
pub enum Decl {
    Var(DimItem),
    Const(Name, Expr),
}

#[derive(Clone, Debug, Default)]
pub struct Module {
    pub name: String,
    pub procs: Vec<Proc>,
    pub decls: Vec<Decl>,
}
