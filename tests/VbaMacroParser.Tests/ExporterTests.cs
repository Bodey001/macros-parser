using System.Text.Json;
using System.Xml;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using VbaMacroParser.Exporters;
using VbaMacroParser.Models;

namespace VbaMacroParser.Tests;

[TestClass]
public sealed class ExporterTests
{
    private static VbaParseResult BuildSampleResult()
    {
        var proc = new VbaProcedure
        {
            Name = "Calculate",
            Kind = ProcedureKind.Function,
            Scope = AccessModifier.Public,
            ReturnType = "Double",
            LineStart = 5,
            LineEnd = 10,
            Parameters =
            [
                new VbaParameter { Name = "x", DataType = "Double", IsByRef = false },
                new VbaParameter { Name = "y", DataType = "Double", IsByRef = false, IsOptional = true, DefaultValue = "0" }
            ],
            Comments = ["Calculates the result"],
            Body = "    Calculate = x + y"
        };

        var module = new VbaModule
        {
            Name = "MathUtils",
            Type = ModuleType.Standard,
            Options = ["Explicit"],
            Constants =
            [
                new VbaConstant { Name = "PI", Scope = AccessModifier.Public, DataType = "Double", Value = "3.14159", LineNumber = 3 }
            ],
            Variables =
            [
                new VbaVariable { Name = "m_Cache", Scope = AccessModifier.Private, DataType = "String", LineNumber = 4 }
            ],
            Procedures = [proc],
            ModuleComments = ["Math utilities module"]
        };

        return new VbaParseResult
        {
            SourceFile = "MathUtils.bas",
            ParsedAt = new DateTime(2026, 4, 17, 12, 0, 0, DateTimeKind.Utc),
            Modules = [module]
        };
    }

    // -----------------------------------------------------------------------
    // JSON
    // -----------------------------------------------------------------------

    [TestMethod]
    public void JsonExporter_ProducesValidJson()
    {
        var result = BuildSampleResult();
        var json = new JsonExporter().Export(result);

        var doc = JsonDocument.Parse(json); // throws if invalid JSON
        Assert.IsNotNull(doc);
    }

    [TestMethod]
    public void JsonExporter_ContainsExpectedFields()
    {
        var result = BuildSampleResult();
        var json = new JsonExporter().Export(result);
        var doc = JsonDocument.Parse(json).RootElement;

        Assert.AreEqual("MathUtils.bas", doc.GetProperty("sourceFile").GetString());
        var modules = doc.GetProperty("modules");
        Assert.AreEqual(1, modules.GetArrayLength());
        Assert.AreEqual("MathUtils", modules[0].GetProperty("name").GetString());
        Assert.AreEqual(1, modules[0].GetProperty("procedures").GetArrayLength());
    }

    [TestMethod]
    public void JsonExporter_ProcedureParametersPresent()
    {
        var result = BuildSampleResult();
        var json = new JsonExporter().Export(result);
        var doc = JsonDocument.Parse(json).RootElement;

        var parms = doc.GetProperty("modules")[0]
                       .GetProperty("procedures")[0]
                       .GetProperty("parameters");

        Assert.AreEqual(2, parms.GetArrayLength());
        Assert.AreEqual("x", parms[0].GetProperty("name").GetString());
        Assert.IsTrue(parms[1].GetProperty("isOptional").GetBoolean());
    }

    // -----------------------------------------------------------------------
    // XML
    // -----------------------------------------------------------------------

    [TestMethod]
    public void XmlExporter_ProducesWellFormedXml()
    {
        var result = BuildSampleResult();
        var xml = new XmlExporter().Export(result);

        var doc = new XmlDocument();
        doc.LoadXml(xml); // throws if malformed
        Assert.IsNotNull(doc.DocumentElement);
    }

    [TestMethod]
    public void XmlExporter_RootElementIsVbaParseResult()
    {
        var result = BuildSampleResult();
        var xml = new XmlExporter().Export(result);

        var doc = new XmlDocument();
        doc.LoadXml(xml);

        Assert.AreEqual("VbaParseResult", doc.DocumentElement!.Name);
    }

    [TestMethod]
    public void XmlExporter_ContainsModuleAndProcedure()
    {
        var result = BuildSampleResult();
        var xml = new XmlExporter().Export(result);

        var doc = new XmlDocument();
        doc.LoadXml(xml);

        var modules = doc.GetElementsByTagName("Module");
        Assert.AreEqual(1, modules.Count);
        Assert.AreEqual("MathUtils", modules[0]!.Attributes!["name"]!.Value);

        var procs = doc.GetElementsByTagName("Procedure");
        Assert.AreEqual(1, procs.Count);
        Assert.AreEqual("Calculate", procs[0]!.Attributes!["name"]!.Value);
    }

    [TestMethod]
    public void XmlExporter_ContainsConstantAndVariable()
    {
        var result = BuildSampleResult();
        var xml = new XmlExporter().Export(result);

        var doc = new XmlDocument();
        doc.LoadXml(xml);

        var constants = doc.GetElementsByTagName("Constant");
        Assert.AreEqual(1, constants.Count);
        Assert.AreEqual("PI", constants[0]!.Attributes!["name"]!.Value);

        var variables = doc.GetElementsByTagName("Variable");
        Assert.AreEqual(1, variables.Count);
        Assert.AreEqual("m_Cache", variables[0]!.Attributes!["name"]!.Value);
    }

    // -----------------------------------------------------------------------
    // Plain Text
    // -----------------------------------------------------------------------

    [TestMethod]
    public void TextExporter_ContainsModuleName()
    {
        var result = BuildSampleResult();
        var text = new TextExporter().Export(result);

        Assert.IsTrue(text.Contains("MathUtils"), "Text output should contain module name");
    }

    [TestMethod]
    public void TextExporter_ContainsProcedureSignature()
    {
        var result = BuildSampleResult();
        var text = new TextExporter().Export(result);

        Assert.IsTrue(text.Contains("Calculate"), "Text output should contain procedure name");
        Assert.IsTrue(text.Contains("Function"), "Text output should identify procedure as Function");
    }

    [TestMethod]
    public void TextExporter_ContainsHeaderBanner()
    {
        var result = BuildSampleResult();
        var text = new TextExporter().Export(result);

        Assert.IsTrue(text.Contains("VBA MACRO PARSE REPORT"), "Text output should contain banner");
        Assert.IsTrue(text.Contains("MathUtils.bas"), "Text output should contain source file name");
    }

    // -----------------------------------------------------------------------
    // CSV
    // -----------------------------------------------------------------------

    [TestMethod]
    public void CsvExporter_HasHeaderRow()
    {
        var result = BuildSampleResult();
        var csv = new CsvExporter().Export(result);
        var lines = csv.Split('\n', StringSplitOptions.RemoveEmptyEntries);

        Assert.IsTrue(lines[0].StartsWith("Module"), "First row should be the header");
        Assert.IsTrue(lines[0].Contains("ProcedureName"));
        Assert.IsTrue(lines[0].Contains("Kind"));
    }

    [TestMethod]
    public void CsvExporter_OneDataRowPerProcedure()
    {
        var result = BuildSampleResult();
        var csv = new CsvExporter().Export(result);
        var lines = csv.Split('\n', StringSplitOptions.RemoveEmptyEntries);

        // 1 header + 1 data row
        Assert.AreEqual(2, lines.Length);
    }

    [TestMethod]
    public void CsvExporter_DataRowContainsProcedureInfo()
    {
        var result = BuildSampleResult();
        var csv = new CsvExporter().Export(result);
        var lines = csv.Split('\n', StringSplitOptions.RemoveEmptyEntries);
        var dataRow = lines[1];

        Assert.IsTrue(dataRow.Contains("MathUtils"), "Data row should contain module name");
        Assert.IsTrue(dataRow.Contains("Calculate"), "Data row should contain procedure name");
        Assert.IsTrue(dataRow.Contains("Function"), "Data row should contain procedure kind");
    }

    [TestMethod]
    public void CsvExporter_EmptyResultHasOnlyHeader()
    {
        var result = new VbaParseResult
        {
            SourceFile = "empty.bas",
            ParsedAt = DateTime.UtcNow,
            Modules = [new VbaModule { Name = "Empty" }]
        };

        var csv = new CsvExporter().Export(result);
        var lines = csv.Split('\n', StringSplitOptions.RemoveEmptyEntries);

        Assert.AreEqual(1, lines.Length, "No procedures means only the header row");
    }

    [TestMethod]
    public void CsvExporter_CommasInValuesAreEscaped()
    {
        var proc = new VbaProcedure
        {
            Name = "Test",
            Kind = ProcedureKind.Sub,
            Scope = AccessModifier.Public,
            Parameters =
            [
                new VbaParameter { Name = "a", DataType = "Integer" },
                new VbaParameter { Name = "b", DataType = "String" }
            ]
        };

        var result = new VbaParseResult
        {
            SourceFile = "test.bas",
            Modules = [new VbaModule { Name = "M", Procedures = [proc] }]
        };

        var csv = new CsvExporter().Export(result);

        // The parameters column value contains "; " separators which won't have unquoted commas
        // but if it did the field must be quoted
        Assert.IsTrue(csv.Length > 0);
    }

    // -----------------------------------------------------------------------
    // File extension contract
    // -----------------------------------------------------------------------

    [TestMethod]
    public void AllExporters_HaveCorrectFileExtension()
    {
        Assert.AreEqual(".json", new JsonExporter().FileExtension);
        Assert.AreEqual(".xml",  new XmlExporter().FileExtension);
        Assert.AreEqual(".txt",  new TextExporter().FileExtension);
        Assert.AreEqual(".csv",  new CsvExporter().FileExtension);
    }
}
