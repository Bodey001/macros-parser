using System.Text.Json.Serialization;

namespace VbaMacroParser.Models;

public sealed class VbaConstant
{
    [JsonPropertyName("name")]
    public string Name { get; set; } = string.Empty;

    [JsonPropertyName("scope")]
    [JsonConverter(typeof(JsonStringEnumConverter))]
    public AccessModifier Scope { get; set; }

    [JsonPropertyName("dataType")]
    public string DataType { get; set; } = "Variant";

    [JsonPropertyName("value")]
    public string Value { get; set; } = string.Empty;

    [JsonPropertyName("lineNumber")]
    public int LineNumber { get; set; }
}
