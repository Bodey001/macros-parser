using VbaMacroParser.Models;

namespace VbaMacroParser.Exporters;

public interface IExporter
{
    /// <summary>File extension produced by this exporter, including leading dot.</summary>
    string FileExtension { get; }

    /// <summary>Serialises the parse result to a string.</summary>
    string Export(VbaParseResult result);
}
