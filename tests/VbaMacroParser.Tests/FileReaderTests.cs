using System.Text;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using VbaMacroParser.IO;

namespace VbaMacroParser.Tests;

[TestClass]
public sealed class FileReaderTests
{
    private string _tempDir = null!;

    [TestInitialize]
    public void Setup()
    {
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
        _tempDir = Path.Combine(Path.GetTempPath(), "VbaParserTests_" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(_tempDir);
    }

    [TestCleanup]
    public void Teardown()
    {
        if (Directory.Exists(_tempDir))
            Directory.Delete(_tempDir, recursive: true);
    }

    // -----------------------------------------------------------------------
    // Encoding detection
    // -----------------------------------------------------------------------

    [TestMethod]
    public void ReadLines_Utf8WithBom_ReadsCorrectly()
    {
        var path = WriteTempFile("utf8_bom.bas", "Sub Foo()\r\nEnd Sub\r\n", new UTF8Encoding(encoderShouldEmitUTF8Identifier: true));
        var lines = FileReader.ReadLines(path);

        Assert.IsTrue(lines.Length >= 2);
        Assert.IsTrue(lines[0].Contains("Sub Foo"), $"Expected 'Sub Foo' but got '{lines[0]}'");
    }

    [TestMethod]
    public void ReadLines_Utf8WithoutBom_ReadsCorrectly()
    {
        var path = WriteTempFile("utf8_no_bom.bas", "Sub Foo()\r\nEnd Sub\r\n", new UTF8Encoding(false));
        var lines = FileReader.ReadLines(path);

        Assert.IsTrue(lines.Any(l => l.Contains("Sub Foo")));
    }

    [TestMethod]
    public void ReadLines_Utf16LittleEndian_ReadsCorrectly()
    {
        var path = WriteTempFile("utf16.bas", "Sub Foo()\r\nEnd Sub\r\n", new UnicodeEncoding(bigEndian: false, byteOrderMark: true));
        var lines = FileReader.ReadLines(path);

        Assert.IsTrue(lines.Any(l => l.Contains("Sub Foo")));
    }

    [TestMethod]
    public void ReadLines_Ansi_ReadsCorrectly()
    {
        // Windows-1252 ANSI file — no BOM
        var encoding = Encoding.GetEncoding(1252);
        var path = WriteTempFile("ansi.bas", "Sub Grüßen()\r\nEnd Sub\r\n", encoding);
        var lines = FileReader.ReadLines(path);

        Assert.IsTrue(lines.Length >= 1);
    }

    // -----------------------------------------------------------------------
    // Unsupported formats
    // -----------------------------------------------------------------------

    [TestMethod]
    public void ReadLines_XlsmFile_ThrowsUnsupportedFormatException()
    {
        var path = Path.Combine(_tempDir, "workbook.xlsm");
        File.WriteAllBytes(path, [0x50, 0x4B, 0x03, 0x04]); // ZIP magic bytes
        Assert.ThrowsExactly<UnsupportedFormatException>(() => FileReader.ReadLines(path));
    }

    [TestMethod]
    public void ReadLines_DocmFile_ThrowsUnsupportedFormatException()
    {
        var path = Path.Combine(_tempDir, "document.docm");
        File.WriteAllBytes(path, [0xD0, 0xCF, 0x11, 0xE0]); // CFB magic bytes
        Assert.ThrowsExactly<UnsupportedFormatException>(() => FileReader.ReadLines(path));
    }

    [TestMethod]
    public void ReadLines_XlsbFile_ThrowsUnsupportedFormatException()
    {
        var path = Path.Combine(_tempDir, "workbook.xlsb");
        File.WriteAllBytes(path, [0x50, 0x4B, 0x03, 0x04]);
        Assert.ThrowsExactly<UnsupportedFormatException>(() => FileReader.ReadLines(path));
    }

    // -----------------------------------------------------------------------
    // File not found
    // -----------------------------------------------------------------------

    [TestMethod]
    public void ReadLines_MissingFile_ThrowsFileNotFoundException()
    {
        Assert.ThrowsExactly<FileNotFoundException>(() => FileReader.ReadLines(Path.Combine(_tempDir, "does_not_exist.bas")));
    }

    // -----------------------------------------------------------------------
    // IsKnownTextFormat
    // -----------------------------------------------------------------------

    [TestMethod]
    public void IsKnownTextFormat_KnownExtensions_ReturnsTrue()
    {
        Assert.IsTrue(FileReader.IsKnownTextFormat("module.bas"));
        Assert.IsTrue(FileReader.IsKnownTextFormat("MyClass.cls"));
        Assert.IsTrue(FileReader.IsKnownTextFormat("Form1.frm"));
        Assert.IsTrue(FileReader.IsKnownTextFormat("code.vba"));
        Assert.IsTrue(FileReader.IsKnownTextFormat("script.txt"));
        Assert.IsTrue(FileReader.IsKnownTextFormat("Module.vb"));
    }

    [TestMethod]
    public void IsKnownTextFormat_BinaryExtensions_ReturnsFalse()
    {
        Assert.IsFalse(FileReader.IsKnownTextFormat("Book.xlsm"));
        Assert.IsFalse(FileReader.IsKnownTextFormat("Doc.docm"));
        Assert.IsFalse(FileReader.IsKnownTextFormat("Sheet.xlsb"));
    }

    // -----------------------------------------------------------------------
    // DetectEncoding
    // -----------------------------------------------------------------------

    [TestMethod]
    public void DetectEncoding_Utf8Bom_ReturnsUtf8()
    {
        var path = WriteTempFile("bom.bas", "content", new UTF8Encoding(true));
        var enc = FileReader.DetectEncoding(path);
        Assert.IsInstanceOfType(enc, typeof(UTF8Encoding));
    }

    [TestMethod]
    public void DetectEncoding_Utf16Le_ReturnsUnicodeEncoding()
    {
        var path = WriteTempFile("utf16le.bas", "Sub X()", new UnicodeEncoding(false, true));
        var enc = FileReader.DetectEncoding(path);
        Assert.IsInstanceOfType(enc, typeof(UnicodeEncoding));
    }

    // -----------------------------------------------------------------------
    // Helpers
    // -----------------------------------------------------------------------

    private string WriteTempFile(string name, string content, Encoding encoding)
    {
        var path = Path.Combine(_tempDir, name);
        File.WriteAllText(path, content, encoding);
        return path;
    }
}
