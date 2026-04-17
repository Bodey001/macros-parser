using System.Text;
using VbaMacroParser.Models;

namespace VbaMacroParser.Exporters;

/// <summary>RFC-4180 compliant CSV exporter. One row per procedure.</summary>
public sealed class CsvExporter : IExporter
{
    public string FileExtension => ".csv";

    private static readonly string[] Headers =
    [
        "Module", "ModuleType", "ProcedureName", "Kind", "Scope",
        "IsStatic", "ReturnType", "Parameters", "LineStart", "LineEnd",
        "CommentCount", "BodyLineCount"
    ];

    public string Export(VbaParseResult result)
    {
        var sb = new StringBuilder();
        sb.AppendLine(string.Join(",", Headers));

        foreach (var module in result.Modules)
        {
            foreach (var proc in module.Procedures)
            {
                var paramStr = string.Join("; ", proc.Parameters.Select(p =>
                {
                    var pass = p.IsByRef ? "ByRef" : "ByVal";
                    var opt = p.IsOptional ? "Optional " : string.Empty;
                    return $"{opt}{pass} {p.Name} As {p.DataType}";
                }));

                sb.AppendLine(string.Join(",",
                    Escape(module.Name),
                    Escape(module.Type.ToString()),
                    Escape(proc.Name),
                    Escape(proc.Kind.ToString()),
                    Escape(proc.Scope.ToString()),
                    Escape(proc.IsStatic.ToString()),
                    Escape(proc.ReturnType ?? string.Empty),
                    Escape(paramStr),
                    proc.LineStart.ToString(),
                    proc.LineEnd.ToString(),
                    proc.Comments.Count.ToString(),
                    CountLines(proc.Body).ToString()
                ));
            }
        }

        return sb.ToString();
    }

    private static string Escape(string value)
    {
        if (value.Contains(',') || value.Contains('"') || value.Contains('\n') || value.Contains('\r'))
            return "\"" + value.Replace("\"", "\"\"") + "\"";
        return value;
    }

    private static int CountLines(string body)
    {
        if (string.IsNullOrEmpty(body)) return 0;
        return body.Split('\n').Length;
    }
}
