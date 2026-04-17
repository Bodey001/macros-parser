using System.Text;
using VbaMacroParser.Models;

namespace VbaMacroParser.Exporters;

public sealed class TextExporter : IExporter
{
    public string FileExtension => ".txt";

    public string Export(VbaParseResult result)
    {
        var sb = new StringBuilder();

        sb.AppendLine("=" + new string('=', 78));
        sb.AppendLine("VBA MACRO PARSE REPORT");
        sb.AppendLine($"Source File : {result.SourceFile}");
        sb.AppendLine($"Parsed At   : {result.ParsedAt:yyyy-MM-dd HH:mm:ss} UTC");
        sb.AppendLine($"Modules     : {result.Modules.Count}");
        sb.AppendLine("=" + new string('=', 78));

        foreach (var module in result.Modules)
            WriteModule(sb, module);

        return sb.ToString();
    }

    private static void WriteModule(StringBuilder sb, VbaModule module)
    {
        sb.AppendLine();
        sb.AppendLine($"MODULE: {module.Name}  [{module.Type}]");
        sb.AppendLine(new string('-', 60));

        if (module.Options.Count > 0)
        {
            sb.AppendLine("  Options:");
            foreach (var opt in module.Options)
                sb.AppendLine($"    Option {opt}");
        }

        if (module.ModuleComments.Count > 0)
        {
            sb.AppendLine("  Module Comments:");
            foreach (var c in module.ModuleComments)
                sb.AppendLine($"    ' {c}");
        }

        if (module.Constants.Count > 0)
        {
            sb.AppendLine();
            sb.AppendLine("  Constants:");
            foreach (var c in module.Constants)
            {
                var scope = c.Scope == AccessModifier.Default ? string.Empty : c.Scope + " ";
                sb.AppendLine($"    {scope}Const {c.Name} As {c.DataType} = {c.Value}  [line {c.LineNumber}]");
            }
        }

        if (module.Variables.Count > 0)
        {
            sb.AppendLine();
            sb.AppendLine("  Variables:");
            foreach (var v in module.Variables)
            {
                var scope = v.Scope == AccessModifier.Default ? "Dim" : v.Scope.ToString();
                var arrStr = v.IsArray ? $"({v.ArrayBounds})" : string.Empty;
                sb.AppendLine($"    {scope} {v.Name}{arrStr} As {v.DataType}  [line {v.LineNumber}]");
            }
        }

        if (module.Procedures.Count > 0)
        {
            sb.AppendLine();
            sb.AppendLine("  Procedures:");
            foreach (var proc in module.Procedures)
                WriteProcedure(sb, proc);
        }
    }

    private static void WriteProcedure(StringBuilder sb, VbaProcedure proc)
    {
        sb.AppendLine();
        sb.AppendLine($"    {proc.Signature}");
        sb.AppendLine($"      Lines  : {proc.LineStart} - {proc.LineEnd}");
        sb.AppendLine($"      Kind   : {proc.Kind}");
        sb.AppendLine($"      Scope  : {proc.Scope}");

        if (proc.Parameters.Count > 0)
        {
            sb.AppendLine("      Parameters:");
            foreach (var p in proc.Parameters)
            {
                var pass = p.IsByRef ? "ByRef" : "ByVal";
                var opt = p.IsOptional ? "Optional " : string.Empty;
                var arr = p.IsParamArray ? "ParamArray " : string.Empty;
                var def = p.DefaultValue != null ? $" = {p.DefaultValue}" : string.Empty;
                sb.AppendLine($"        {opt}{arr}{pass} {p.Name} As {p.DataType}{def}");
            }
        }

        if (proc.Comments.Count > 0)
        {
            sb.AppendLine("      Comments:");
            foreach (var c in proc.Comments)
                sb.AppendLine($"        ' {c}");
        }
    }
}
