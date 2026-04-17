using System.Text.Json.Serialization;

namespace VbaMacroParser.Models;

public sealed class VbaVariable
{
    [JsonPropertyName("name")]
    public string Name { get; set; } = string.Empty;

    [JsonPropertyName("scope")]
    [JsonConverter(typeof(JsonStringEnumConverter))]
    public AccessModifier Scope { get; set; }

    [JsonPropertyName("dataType")]
    public string DataType { get; set; } = "Variant";

    [JsonPropertyName("isArray")]
    public bool IsArray { get; set; }

    [JsonPropertyName("arrayBounds")]
    public string? ArrayBounds { get; set; }

    [JsonPropertyName("isStatic")]
    public bool IsStatic { get; set; }

    [JsonPropertyName("lineNumber")]
    public int LineNumber { get; set; }
}
