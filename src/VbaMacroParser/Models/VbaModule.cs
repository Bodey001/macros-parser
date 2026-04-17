using System.Text.Json.Serialization;

namespace VbaMacroParser.Models;

public sealed class VbaModule
{
    [JsonPropertyName("name")]
    public string Name { get; set; } = string.Empty;

    [JsonPropertyName("type")]
    [JsonConverter(typeof(JsonStringEnumConverter))]
    public ModuleType Type { get; set; }

    [JsonPropertyName("options")]
    public List<string> Options { get; set; } = [];

    [JsonPropertyName("constants")]
    public List<VbaConstant> Constants { get; set; } = [];

    [JsonPropertyName("variables")]
    public List<VbaVariable> Variables { get; set; } = [];

    [JsonPropertyName("procedures")]
    public List<VbaProcedure> Procedures { get; set; } = [];

    [JsonPropertyName("moduleComments")]
    public List<string> ModuleComments { get; set; } = [];
}
