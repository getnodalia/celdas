//! Whole VBA programs run through the interpreter, checked by what they `Debug.Print`.
#![allow(clippy::unwrap_used, clippy::expect_used, clippy::panic)]

use celdas_vba::{ModuleKind, NoHost, Project, SourceModule, Value, Vm};

fn project(src: &str) -> Project {
    Project::from_sources(vec![SourceModule { name: "Module1".into(), kind: ModuleKind::Standard, source: src.into() }])
}

/// Runs `Main` and returns what it printed.
fn run(src: &str) -> Vec<String> {
    let p = project(src);
    let mut host = NoHost;
    let mut vm = Vm::new(&p, &mut host);
    vm.run("Main", Vec::new()).unwrap_or_else(|e| panic!("{e}"));
    vm.debug
}

fn fails(src: &str) -> String {
    let p = project(src);
    let mut host = NoHost;
    let mut vm = Vm::new(&p, &mut host);
    vm.run("Main", Vec::new()).expect_err("should fail").to_string()
}

#[test]
fn arithmetic_and_precedence() {
    let out = run(
        "Sub Main()\nDebug.Print 2 + 3 * 4\nDebug.Print -2 ^ 2\nDebug.Print 7 \\ 2, 7 Mod 3, 7 / 2\nDebug.Print \"a\" & 1 + 2\nDebug.Print Not 5 = 5\nEnd Sub",
    );
    assert_eq!(out, ["14", "-4", "3\t1\t3.5", "a3", "False"]);
}

#[test]
fn loops_and_conditions() {
    let out = run(
        "Sub Main()\nDim i As Long, s As Long\nFor i = 1 To 10 Step 3\ns = s + i\nNext i\nDebug.Print s\nDo While s > 0\ns = s - 10\nLoop\nDebug.Print s\nFor i = 5 To 1 Step -2\nDebug.Print i\nNext\nIf s < 0 Then Debug.Print \"neg\" Else Debug.Print \"pos\"\nEnd Sub",
    );
    assert_eq!(out[0], "22");
    assert_eq!(out[1], "-8");
    assert!(out.contains(&"neg".to_string()));
}

#[test]
fn select_case_and_strings() {
    let out = run(
        "Function K(n)\nSelect Case n\nCase 1, 2\nK = \"low\"\nCase 3 To 9\nK = \"mid\"\nCase Is >= 10\nK = \"high\"\nCase Else\nK = \"?\"\nEnd Select\nEnd Function\nSub Main()\nDebug.Print K(2), K(5), K(12), K(0)\nDebug.Print UCase(Left(\"hello\", 3)) & Mid(\"hello\", 2, 2) & Right(\"hello\", 1)\nDebug.Print InStr(\"abcabc\", \"c\"), InStrRev(\"abcabc\", \"c\"), Len(\"abc\")\nDebug.Print Replace(\"a-b-c\", \"-\", \"+\"), Trim(\"  x  \")\nEnd Sub",
    );
    assert_eq!(out[0], "low\tmid\thigh\t?");
    assert_eq!(out[1], "HELelo");
    assert_eq!(out[2], "3\t6\t3");
    assert_eq!(out[3], "a+b+c\tx");
}

#[test]
fn arrays_split_join_and_redim() {
    let out = run(
        "Sub Main()\nDim a(1 To 3) As Double\nDim i\nFor i = 1 To 3\na(i) = i * 1.5\nNext\nDebug.Print a(1) + a(2) + a(3), UBound(a), LBound(a)\nDim v\nv = Split(\"x,y,z\", \",\")\nDebug.Print UBound(v), Join(v, \"-\")\nReDim Preserve a(1 To 5)\nDebug.Print a(3), UBound(a)\nDim m(1 To 2, 1 To 2)\nm(2, 1) = 7\nDebug.Print m(2, 1)\nEnd Sub",
    );
    assert_eq!(out, ["9\t3\t1", "2\tx-y-z", "4.5\t5", "7"]);
}

#[test]
fn byref_byval_and_recursion() {
    let out = run(
        "Sub Bump(ByRef n As Long, ByVal m As Long)\nn = n + 1\nm = m + 1\nEnd Sub\nFunction Fact(n As Long) As Long\nIf n <= 1 Then Fact = 1 Else Fact = n * Fact(n - 1)\nEnd Function\nSub Main()\nDim a As Long, b As Long\na = 1: b = 1\nBump a, b\nDebug.Print a, b, Fact(10)\nEnd Sub",
    );
    assert_eq!(out, ["2\t1\t3628800"]);
}

#[test]
fn error_handling_goto_and_resume_next() {
    let out = run(
        "Sub Main()\nOn Error GoTo Oops\nDim x As Long\nx = 1 / 0\nDebug.Print \"skipped\"\nGoTo Done\nOops:\nDebug.Print \"caught \" & Err.Number & \" \" & Err.Description\nDone:\nOn Error Resume Next\nx = CLng(\"abc\")\nDebug.Print Err.Number\nErr.Clear\nDebug.Print Err.Number\nEnd Sub",
    );
    assert!(out[0].starts_with("caught 11 "), "{out:?}");
    assert_eq!(out[1], "13");
    assert_eq!(out[2], "0");
}

#[test]
fn collections_and_dictionaries() {
    let out = run(
        "Sub Main()\nDim c As New Collection, d As Object, v\nc.Add 10\nc.Add 20, \"k\"\nDebug.Print c.Count, c(\"k\"), c(1)\nFor Each v In c\nDebug.Print v\nNext\nSet d = CreateObject(\"Scripting.Dictionary\")\nd.Add \"a\", 1\nd(\"b\") = 2\nDebug.Print d.Count, d.Exists(\"b\"), d(\"a\")\nEnd Sub",
    );
    assert_eq!(out[0], "2\t20\t10");
    assert!(out.iter().any(|l| l == "2\tTrue\t1"), "{out:?}");
}

#[test]
fn format_and_conversions() {
    let out = run(
        "Sub Main()\nDebug.Print Format(1234.5, \"#,##0.00\"), Format(0.256, \"0.0%\")\nDebug.Print CInt(2.5), CInt(3.5), Round(2.675, 2), Int(-2.5), Fix(-2.5)\nDebug.Print CStr(12) & \"!\", CDbl(\"3.25\") * 2, Val(\"12abc\")\nDebug.Print TypeName(1), TypeName(1.5), TypeName(\"s\"), TypeName(Empty)\nEnd Sub",
    );
    assert_eq!(out[0], "1,234.50\t25.6%");
    assert_eq!(out[1], "2\t4\t2.68\t-3\t-2");
    assert_eq!(out[2], "12!\t6.5\t12");
    assert_eq!(out[3], "Integer\tDouble\tString\tEmpty");
}

#[test]
fn constants_with_and_exit() {
    let out = run(
        "Const N = 4\nPublic total As Long\nSub Add(k)\ntotal = total + k\nEnd Sub\nSub Main()\nDim i\nFor i = 1 To 100\nIf i > N Then Exit For\nAdd i\nNext\nDebug.Print total, i\nEnd Sub",
    );
    assert_eq!(out, ["10\t5"]);
}

#[test]
fn runaway_loops_are_stopped() {
    let p = project("Sub Main()\nDo\nLoop\nEnd Sub");
    let mut host = NoHost;
    let mut vm = Vm::new(&p, &mut host);
    vm.set_max_steps(10_000);
    let e = vm.run("Main", Vec::new()).expect_err("must stop");
    assert!(e.to_string().contains("too long"), "{e}");
}

#[test]
fn recursion_is_bounded() {
    assert!(fails("Sub Main()\nMain\nEnd Sub").contains("stack"));
}

#[test]
fn missing_procedures_and_syntax_errors_are_reported() {
    let p = project("Sub Main()\nEnd Sub");
    let mut host = NoHost;
    let mut vm = Vm::new(&p, &mut host);
    assert!(vm.run("Nope", Vec::new()).is_err());
    let bad = project("Sub Main(\nx = = 1\nEnd Sub");
    let mut host = NoHost;
    let mut vm = Vm::new(&bad, &mut host);
    assert!(vm.run("Main", Vec::new()).is_err());
}

#[test]
fn functions_return_values_to_the_caller() {
    let p = project("Function Twice(x)\nTwice = x * 2\nEnd Function");
    let mut host = NoHost;
    let mut vm = Vm::new(&p, &mut host);
    assert_eq!(vm.run("Twice", vec![Value::Long(21)]).unwrap(), Value::Double(42.0));
}

/// A host with one object, `Box`, that remembers what was assigned to its `Item` property.
mod assigning_to_host_members {
    use celdas_vba::{CallArg, Host, HostObject, Obj, Project, Value, VbError, Vm};
    use std::any::Any;

    struct Thing(&'static str);

    impl HostObject for Thing {
        fn class(&self) -> String {
            "Thing".into()
        }
        fn as_any(&self) -> &dyn Any {
            self
        }
        fn same(&self, other: &dyn HostObject) -> bool {
            other.as_any().downcast_ref::<Thing>().is_some_and(|o| o.0 == self.0)
        }
    }

    #[derive(Default)]
    struct Probe {
        assigned: Vec<String>,
    }

    impl Host for Probe {
        fn member(&mut self, obj: &Obj, name: &str, _: &[CallArg], assign: Option<Value>) -> Result<Value, VbError> {
            match (obj.downcast::<Thing>().map(|t| t.0), name, assign) {
                (Some("box"), "item", Some(v)) => {
                    self.assigned.push(match v {
                        Value::Object(Some(o)) => format!("object {}", o.class()),
                        other => format!("value {other:?}"),
                    });
                    Ok(Value::Empty)
                }
                (Some("box"), "peer", None) => Ok(Value::Object(Some(Obj::new(Thing("peer"))))),
                // A default member: reading a Thing as a plain value gives its name.
                (Some(n), "", None) => Ok(Value::string(n.to_string())),
                _ => Err(VbError::no_member(name)),
            }
        }

        fn global(&mut self, name: &str, _: &[CallArg], _: Option<Value>) -> Result<Option<Value>, VbError> {
            Ok((name == "box").then(|| Value::Object(Some(Obj::new(Thing("box"))))))
        }
    }

    #[test]
    fn a_property_set_to_an_object_receives_the_object_not_its_default_value() {
        let project = Project::from_sources(vec![celdas_vba::SourceModule {
            name: "Module1".into(),
            kind: celdas_vba::ModuleKind::Standard,
            source: "Sub Main()\nbox.Item = box.Peer\nDim s As String\ns = box.Peer\nbox.Item = s\nEnd Sub\n".into(),
        }]);
        let mut host = Probe::default();
        Vm::new(&project, &mut host).run("Main", Vec::new()).unwrap();
        assert_eq!(host.assigned, ["object Thing", "value Str(\"peer\")"].map(String::from));
    }
}
