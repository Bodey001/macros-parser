using System.Text.Json.Serialization;

namespace VbaMacroParser.Models;

public sealed class VbaParameter
{
    [JsonPropertyName("name")]
    public string Name { get; set; } = string.Empty;

    [JsonPropertyName("dataType")]
    public string DataType { get; set; } = "Variant";

    [JsonPropertyName("isByRef")]
    public bool IsByRef { get; set; } = true;

    [JsonPropertyName("isOptional")]
    public bool IsOptional { get; set; }

    [JsonPropertyName("defaultValue")]
    public string? DefaultValue { get; set; }

    [JsonPropertyName("isParamArray")]
    public bool IsParamArray { get; set; }
}
