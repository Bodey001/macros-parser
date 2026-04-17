using System.Text;

namespace VbaMacroParser.IO;

/// <summary>
/// Reads a text-based VBA source file, auto-detecting encoding via BOM.
/// Throws <see cref="UnsupportedFormatException"/> for known binary Office formats.
/// </summary>
public static class FileReader
{
    private static readonly HashSet<string> BinaryOfficeExtensions =
        new(StringComparer.OrdinalIgnoreCase) { ".xlsm", ".xlsb", ".xls", ".docm", ".docb", ".pptm" };

    private static readonly HashSet<string> SupportedExtensions =
        new(StringComparer.OrdinalIgnoreCase) { ".bas", ".cls", ".frm", ".vba", ".txt", ".vb" };

    public static string[] ReadLines(string filePath)
    {
        if (!File.Exists(filePath))
            throw new FileNotFoundException($"Input file not found: '{filePath}'", filePath);

        var ext = Path.GetExtension(filePath);

        if (BinaryOfficeExtensions.Contains(ext))
            throw new UnsupportedFormatException(filePath);

        var encoding = DetectEncoding(filePath);
        return File.ReadAllLines(filePath, encoding);
    }

    /// <summary>
    /// Returns true when the file extension is a known text-based VBA source.
    /// Files with unknown extensions are still attempted (treated as text).
    /// </summary>
    public static bool IsKnownTextFormat(string filePath) =>
        SupportedExtensions.Contains(Path.GetExtension(filePath));

    /// <summary>
    /// Reads the BOM (if present) to determine encoding; falls back to UTF-8.
    /// For Windows ANSI files without BOM, Windows-1252 is used via a detected codepage.
    /// </summary>
    public static Encoding DetectEncoding(string filePath)
    {
        using var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read);

        Span<byte> bom = stackalloc byte[4];
        var read = fs.Read(bom);

        if (read >= 3 && bom[0] == 0xEF && bom[1] == 0xBB && bom[2] == 0xBF)
            return new UTF8Encoding(encoderShouldEmitUTF8Identifier: true);

        if (read >= 2 && bom[0] == 0xFF && bom[1] == 0xFE)
            return new UnicodeEncoding(bigEndian: false, byteOrderMark: true);

        if (read >= 2 && bom[0] == 0xFE && bom[1] == 0xFF)
            return new UnicodeEncoding(bigEndian: true, byteOrderMark: true);

        if (read >= 4 && bom[0] == 0x00 && bom[1] == 0x00 && bom[2] == 0xFE && bom[3] == 0xFF)
            return new UTF32Encoding(bigEndian: true, byteOrderMark: true);

        // No BOM — heuristic: attempt UTF-8 strict; fall back to Windows-1252 (ANSI).
        try
        {
            fs.Seek(0, SeekOrigin.Begin);
            using var sr = new StreamReader(fs, new UTF8Encoding(false, throwOnInvalidBytes: true), detectEncodingFromByteOrderMarks: false, leaveOpen: true);
            sr.ReadToEnd();
            return new UTF8Encoding(false);
        }
        catch (DecoderFallbackException)
        {
            // Contains bytes invalid in UTF-8 — use Windows-1252 (ANSI)
            return Encoding.GetEncoding(1252);
        }
    }
}
