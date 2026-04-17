using System.Text.Json.Serialization;

namespace VbaMacroParser.Models;

public sealed class VbaProcedure
{
    [JsonPropertyName("name")]
    public string Name { get; set; } = string.Empty;

    [JsonPropertyName("kind")]
    [JsonConverter(typeof(JsonStringEnumConverter))]
    public ProcedureKind Kind { get; set; }

    [JsonPropertyName("scope")]
    [JsonConverter(typeof(JsonStringEnumConverter))]
    public AccessModifier Scope { get; set; }

    [JsonPropertyName("isStatic")]
    public bool IsStatic { get; set; }

    [JsonPropertyName("returnType")]
    public string? ReturnType { get; set; }

    [JsonPropertyName("parameters")]
    public List<VbaParameter> Parameters { get; set; } = [];

    [JsonPropertyName("lineStart")]
    public int LineStart { get; set; }

    [JsonPropertyName("lineEnd")]
    public int LineEnd { get; set; }

    [JsonPropertyName("body")]
    public string Body { get; set; } = string.Empty;

    [JsonPropertyName("comments")]
    public List<string> Comments { get; set; } = [];

    public string Signature
    {
        get
        {
            var scope = Scope == AccessModifier.Default ? string.Empty : Scope.ToString() + " ";
            var staticKw = IsStatic ? "Static " : string.Empty;
            var kindStr = Kind switch
            {
                ProcedureKind.PropertyGet => "Property Get",
                ProcedureKind.PropertyLet => "Property Let",
                ProcedureKind.PropertySet => "Property Set",
                _ => Kind.ToString()
            };
            var parms = string.Join(", ", Parameters.Select(p =>
            {
                var parts = new List<string>();
                if (p.IsOptional) parts.Add("Optional");
                if (p.IsParamArray) parts.Add("ParamArray");
                parts.Add(p.IsByRef ? "ByRef" : "ByVal");
                parts.Add(p.Name);
                parts.Add("As " + p.DataType);
                if (p.DefaultValue != null) parts.Add("= " + p.DefaultValue);
                return string.Join(" ", parts);
            }));
            var ret = ReturnType != null ? $" As {ReturnType}" : string.Empty;
            return $"{scope}{staticKw}{kindStr} {Name}({parms}){ret}";
        }
    }
}
