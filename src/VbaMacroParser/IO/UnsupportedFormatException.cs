namespace VbaMacroParser.IO;

public sealed class UnsupportedFormatException : Exception
{
    public string FilePath { get; }

    public UnsupportedFormatException(string filePath)
        : base($"Binary Office format is not supported without a CFB library: '{filePath}'. " +
               "Extract the VBA source first and re-run with a .bas / .cls / .frm file.")
    {
        FilePath = filePath;
    }
}
