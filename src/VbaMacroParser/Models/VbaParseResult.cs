using System.Text.Json.Serialization;

namespace VbaMacroParser.Models;

public sealed class VbaParseResult
{
    [JsonPropertyName("sourceFile")]
    public string SourceFile { get; set; } = string.Empty;

    [JsonPropertyName("parsedAt")]
    public DateTime ParsedAt { get; set; }

    [JsonPropertyName("modules")]
    public List<VbaModule> Modules { get; set; } = [];

    [JsonPropertyName("parserVersion")]
    public string ParserVersion { get; set; } = "1.0.0";
}
