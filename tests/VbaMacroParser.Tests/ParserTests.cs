using Microsoft.VisualStudio.TestTools.UnitTesting;
using VbaMacroParser.Models;
using VbaMacroParser.Parser;

namespace VbaMacroParser.Tests;

[TestClass]
public sealed class ParserTests
{
    private static VbaParser CreateParser() => new();

    // -----------------------------------------------------------------------
    // Basic Sub and Function detection
    // -----------------------------------------------------------------------

    [TestMethod]
    public void Parse_BasicSubAndFunction_DetectsBothProcedures()
    {
        var lines = """
            Attribute VB_Name = "TestModule"
            Option Explicit

            Public Sub HelloWorld()
                MsgBox "Hello"
            End Sub

            Public Function AddNumbers(ByVal x As Integer, ByVal y As Integer) As Integer
                AddNumbers = x + y
            End Function
            """.Split('\n');

        var result = CreateParser().Parse("TestModule.bas", lines);

        Assert.AreEqual(1, result.Modules.Count);
        var module = result.Modules[0];
        Assert.AreEqual("TestModule", module.Name);
        Assert.AreEqual(2, module.Procedures.Count);

        var sub = module.Procedures[0];
        Assert.AreEqual("HelloWorld", sub.Name);
        Assert.AreEqual(ProcedureKind.Sub, sub.Kind);
        Assert.AreEqual(AccessModifier.Public, sub.Scope);

        var func = module.Procedures[1];
        Assert.AreEqual("AddNumbers", func.Name);
        Assert.AreEqual(ProcedureKind.Function, func.Kind);
        Assert.AreEqual("Integer", func.ReturnType);
        Assert.AreEqual(2, func.Parameters.Count);
    }

    [TestMethod]
    public void Parse_BasicSubAndFunction_LineNumbersAreCorrect()
    {
        var lines = new[]
        {
            "Option Explicit",       // 1
            "",                      // 2
            "Public Sub Foo()",      // 3
            "    ' body",            // 4
            "End Sub",               // 5
        };

        var result = CreateParser().Parse("Foo.bas", lines);
        var proc = result.Modules[0].Procedures[0];

        Assert.AreEqual(3, proc.LineStart);
        Assert.AreEqual(5, proc.LineEnd);
    }

    // -----------------------------------------------------------------------
    // Class module
    // -----------------------------------------------------------------------

    [TestMethod]
    public void Parse_ClassFile_ModuleTypeIsClass()
    {
        var lines = """
            VERSION 1.0 CLASS
            Attribute VB_Name = "MyClass"
            Attribute VB_GlobalNameSpace = False

            Private Sub Class_Initialize()
            End Sub
            """.Split('\n');

        var result = CreateParser().Parse("MyClass.cls", lines);

        Assert.AreEqual(ModuleType.Class, result.Modules[0].Type);
        Assert.AreEqual("MyClass", result.Modules[0].Name);
    }

    // -----------------------------------------------------------------------
    // Options
    // -----------------------------------------------------------------------

    [TestMethod]
    public void Parse_OptionExplicit_IsCollected()
    {
        var lines = new[]
        {
            "Option Explicit",
            "Option Base 1",
            "",
            "Sub Dummy()",
            "End Sub"
        };

        var result = CreateParser().Parse("test.bas", lines);
        var opts = result.Modules[0].Options;

        Assert.AreEqual(2, opts.Count);
        Assert.IsTrue(opts.Contains("Explicit"));
        Assert.IsTrue(opts.Contains("Base 1"));
    }

    // -----------------------------------------------------------------------
    // Comments
    // -----------------------------------------------------------------------

    [TestMethod]
    public void Parse_RemAndApostropheComments_AreCollected()
    {
        var lines = new[]
        {
            "' Module header comment",
            "Rem Another module comment",
            "",
            "Sub Foo()",
            "    ' inline comment",
            "    x = 1 ' trailing comment",
            "End Sub"
        };

        var result = CreateParser().Parse("test.bas", lines);
        var module = result.Modules[0];

        Assert.AreEqual(2, module.ModuleComments.Count);
        Assert.AreEqual("Module header comment", module.ModuleComments[0]);

        var proc = module.Procedures[0];
        Assert.IsTrue(proc.Comments.Count >= 1, "Procedure should have at least the inline comment");
        Assert.IsTrue(proc.Comments.Any(c => c.Contains("inline comment")));
    }

    // -----------------------------------------------------------------------
    // Parameters
    // -----------------------------------------------------------------------

    [TestMethod]
    public void Parse_MultipleParameterKinds_AllParsedCorrectly()
    {
        var lines = new[]
        {
            "Public Sub Complex(ByVal name As String, ByRef count As Integer, Optional ByVal flag As Boolean = False, ParamArray items() As Variant)",
            "End Sub"
        };

        var result = CreateParser().Parse("test.bas", lines);
        var parms = result.Modules[0].Procedures[0].Parameters;

        Assert.AreEqual(4, parms.Count);

        Assert.AreEqual("name", parms[0].Name);
        Assert.IsFalse(parms[0].IsByRef);
        Assert.AreEqual("String", parms[0].DataType);

        Assert.AreEqual("count", parms[1].Name);
        Assert.IsTrue(parms[1].IsByRef);
        Assert.AreEqual("Integer", parms[1].DataType);

        Assert.AreEqual("flag", parms[2].Name);
        Assert.IsTrue(parms[2].IsOptional);
        Assert.IsNotNull(parms[2].DefaultValue);

        Assert.AreEqual("items", parms[3].Name);
        Assert.IsTrue(parms[3].IsParamArray);
    }

    // -----------------------------------------------------------------------
    // Variables and Constants
    // -----------------------------------------------------------------------

    [TestMethod]
    public void Parse_ModuleLevelVariables_AreCollected()
    {
        var lines = new[]
        {
            "Private m_Name As String",
            "Public Counter As Integer",
            "Dim Results() As Double",
            "",
            "Sub Dummy()",
            "End Sub"
        };

        var result = CreateParser().Parse("test.bas", lines);
        var vars = result.Modules[0].Variables;

        Assert.AreEqual(3, vars.Count);
        Assert.AreEqual("m_Name", vars[0].Name);
        Assert.AreEqual(AccessModifier.Private, vars[0].Scope);
        Assert.AreEqual("String", vars[0].DataType);

        Assert.AreEqual("Results", vars[2].Name);
        Assert.IsTrue(vars[2].IsArray);
    }

    [TestMethod]
    public void Parse_Constants_AreCollected()
    {
        var lines = new[]
        {
            "Public Const MAX_SIZE As Integer = 100",
            "Private Const APP_NAME As String = \"MyApp\"",
            "",
            "Sub Dummy()",
            "End Sub"
        };

        var result = CreateParser().Parse("test.bas", lines);
        var consts = result.Modules[0].Constants;

        Assert.AreEqual(2, consts.Count);
        Assert.AreEqual("MAX_SIZE", consts[0].Name);
        Assert.AreEqual("100", consts[0].Value);
        Assert.AreEqual("Integer", consts[0].DataType);
        Assert.AreEqual(AccessModifier.Public, consts[0].Scope);

        Assert.AreEqual("APP_NAME", consts[1].Name);
        Assert.AreEqual(AccessModifier.Private, consts[1].Scope);
    }

    // -----------------------------------------------------------------------
    // Property procedures
    // -----------------------------------------------------------------------

    [TestMethod]
    public void Parse_PropertyGetLetSet_AllDetected()
    {
        var lines = new[]
        {
            "Public Property Get Name() As String",
            "    Name = m_Name",
            "End Property",
            "",
            "Public Property Let Name(ByVal value As String)",
            "    m_Name = value",
            "End Property",
            "",
            "Public Property Set Obj(ByVal value As Object)",
            "    Set m_Obj = value",
            "End Property"
        };

        var result = CreateParser().Parse("MyClass.cls", lines);
        var procs = result.Modules[0].Procedures;

        Assert.AreEqual(3, procs.Count);
        Assert.AreEqual(ProcedureKind.PropertyGet, procs[0].Kind);
        Assert.AreEqual(ProcedureKind.PropertyLet, procs[1].Kind);
        Assert.AreEqual(ProcedureKind.PropertySet, procs[2].Kind);
    }

    // -----------------------------------------------------------------------
    // Edge cases
    // -----------------------------------------------------------------------

    [TestMethod]
    public void Parse_EmptyFile_ReturnsModuleWithNoChildren()
    {
        var result = CreateParser().Parse("empty.bas", Array.Empty<string>());

        Assert.IsNotNull(result);
        Assert.AreEqual(1, result.Modules.Count);
        Assert.AreEqual(0, result.Modules[0].Procedures.Count);
        Assert.AreEqual(0, result.Modules[0].Variables.Count);
        Assert.AreEqual(0, result.Modules[0].Constants.Count);
    }

    [TestMethod]
    public void Parse_FileWithOnlyComments_NoProcedures()
    {
        var lines = new[]
        {
            "' This file contains only comments",
            "' Author: Test",
            "Rem Legacy comment style"
        };

        var result = CreateParser().Parse("comments_only.bas", lines);

        Assert.AreEqual(0, result.Modules[0].Procedures.Count);
        Assert.AreEqual(3, result.Modules[0].ModuleComments.Count);
    }

    [TestMethod]
    public void Parse_PrivateStaticSub_FlagsCorrect()
    {
        var lines = new[]
        {
            "Private Static Sub CachedLoad()",
            "End Sub"
        };

        var result = CreateParser().Parse("test.bas", lines);
        var proc = result.Modules[0].Procedures[0];

        Assert.AreEqual(AccessModifier.Private, proc.Scope);
        Assert.IsTrue(proc.IsStatic);
    }

    [TestMethod]
    public void Parse_ProcedureBody_IsPreserved()
    {
        var lines = new[]
        {
            "Sub DoWork()",
            "    Dim x As Integer",
            "    x = 42",
            "    MsgBox x",
            "End Sub"
        };

        var result = CreateParser().Parse("test.bas", lines);
        var body = result.Modules[0].Procedures[0].Body;

        Assert.IsTrue(body.Contains("x = 42"), "Body should contain the assignment");
        Assert.IsTrue(body.Contains("MsgBox"), "Body should contain the MsgBox call");
    }

    [TestMethod]
    public void Parse_SourceFileAndParsedAt_AreSet()
    {
        var before = DateTime.UtcNow.AddSeconds(-1);
        var result = CreateParser().Parse("path/to/test.bas", ["Sub Foo()", "End Sub"]);
        var after = DateTime.UtcNow.AddSeconds(1);

        Assert.AreEqual("path/to/test.bas", result.SourceFile);
        Assert.IsTrue(result.ParsedAt >= before && result.ParsedAt <= after);
    }
}
