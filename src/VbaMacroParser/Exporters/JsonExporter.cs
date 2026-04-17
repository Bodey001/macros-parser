using System.Text.Json;
using System.Text.Json.Serialization;
using VbaMacroParser.Models;

namespace VbaMacroParser.Exporters;

public sealed class JsonExporter : IExporter
{
    public string FileExtension => ".json";

    private static readonly JsonSerializerOptions Options = new()
    {
        WriteIndented = true,
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
        Converters = { new JsonStringEnumConverter() }
    };

    public string Export(VbaParseResult result) =>
        JsonSerializer.Serialize(result, Options);
}
