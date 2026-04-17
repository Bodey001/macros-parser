using System.Text;
using System.Text.RegularExpressions;
using VbaMacroParser.Models;

namespace VbaMacroParser.Parser;

public sealed class VbaParser : IVbaParser
{
    // Matches individual variable declarations within a Dim/Public/Private statement.
    // Handles:  Name As Type, Name(bounds) As Type, Name
    private static readonly Regex ReVarDecl = new(
        @"(\w+)(\([^)]*\))?\s*(?:As\s+(\w+))?",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);

    // Matches a single parameter: [Optional] [ByRef|ByVal] [ParamArray] Name [()] [As Type] [= Default]
    private static readonly Regex ReParam = new(
        @"(?:(Optional)\s+)?(?:(ParamArray)\s+)?(?:(ByRef|ByVal)\s+)?(\w+)(\(\))?\s*(?:As\s+(\w+(?:\(\))?))?\s*(?:=\s*(.+))?",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);

    public VbaParseResult Parse(string sourceFilePath, IEnumerable<string> lines)
    {
        var result = new VbaParseResult
        {
            SourceFile = sourceFilePath,
            ParsedAt = DateTime.UtcNow
        };

        var module = CreateModule(sourceFilePath);
        result.Modules.Add(module);

        VbaProcedure? currentProc = null;
        var bodyBuilder = new StringBuilder();

        foreach (var token in VbaLexer.Tokenise(lines))
        {
            switch (token.Kind)
            {
                case TokenKind.Attribute:
                    if (currentProc is null)
                        module.Name = token.Groups![1].Value.Trim();
                    break;

                case TokenKind.Option:
                    if (currentProc is null)
                        module.Options.Add(token.Groups![1].Value.Trim());
                    break;

                case TokenKind.Const:
                    if (currentProc is null)
                        module.Constants.Add(ParseConstant(token));
                    else
                    {
                        bodyBuilder.AppendLine(token.Raw);
                        AppendInlineComments(token.Raw, currentProc.Comments);
                    }
                    break;

                case TokenKind.Variable:
                    if (currentProc is null)
                        module.Variables.AddRange(ParseVariables(token));
                    else
                        bodyBuilder.AppendLine(token.Raw);
                    break;

                case TokenKind.ProcedureOpen:
                    currentProc = ParseProcedureOpen(token);
                    bodyBuilder.Clear();
                    break;

                case TokenKind.ProcedureClose:
                    if (currentProc is not null)
                    {
                        currentProc.LineEnd = token.LineNumber;
                        currentProc.Body = bodyBuilder.ToString().TrimEnd();
                        module.Procedures.Add(currentProc);
                        currentProc = null;
                        bodyBuilder.Clear();
                    }
                    break;

                case TokenKind.Comment:
                    var commentText = token.Groups![1].Value.Trim();
                    if (currentProc is not null)
                        currentProc.Comments.Add(commentText);
                    else
                        module.ModuleComments.Add(commentText);
                    break;

                case TokenKind.Other:
                    if (currentProc is not null)
                    {
                        bodyBuilder.AppendLine(token.Raw);
                        AppendInlineComments(token.Raw, currentProc.Comments);
                    }
                    break;

                case TokenKind.Blank:
                    if (currentProc is not null)
                        bodyBuilder.AppendLine(string.Empty);
                    break;
            }
        }

        // Handle unclosed procedure (malformed input — store what we have)
        if (currentProc is not null)
        {
            currentProc.Body = bodyBuilder.ToString().TrimEnd();
            module.Procedures.Add(currentProc);
        }

        return result;
    }

    private static VbaModule CreateModule(string sourceFilePath)
    {
        var ext = Path.GetExtension(sourceFilePath).ToLowerInvariant();
        var moduleType = ext switch
        {
            ".cls" => ModuleType.Class,
            ".frm" => ModuleType.Form,
            _ => ModuleType.Standard
        };

        return new VbaModule
        {
            Name = Path.GetFileNameWithoutExtension(sourceFilePath),
            Type = moduleType
        };
    }

    private static VbaConstant ParseConstant(Token token)
    {
        var g = token.Groups!;
        return new VbaConstant
        {
            Scope = ParseAccessModifier(g[1].Value),
            Name = g[2].Value.Trim(),
            DataType = string.IsNullOrWhiteSpace(g[3].Value) ? "Variant" : g[3].Value.Trim(),
            Value = g[4].Value.Trim(),
            LineNumber = token.LineNumber
        };
    }

    private static IEnumerable<VbaVariable> ParseVariables(Token token)
    {
        var g = token.Groups!;
        var scopeStr = g[1].Value;
        var scope = ParseAccessModifier(scopeStr);
        var isStatic = scopeStr.Equals("Static", StringComparison.OrdinalIgnoreCase);
        var rest = g[2].Value;

        // Multiple declarations: Dim x As Integer, y As String
        var declarations = SplitDeclarations(rest);
        foreach (var decl in declarations)
        {
            var m = ReVarDecl.Match(decl.Trim());
            if (!m.Success) continue;

            var name = m.Groups[1].Value;
            var arrayPart = m.Groups[2].Value;
            var dataType = string.IsNullOrWhiteSpace(m.Groups[3].Value) ? "Variant" : m.Groups[3].Value;

            yield return new VbaVariable
            {
                Name = name,
                Scope = scope,
                DataType = dataType,
                IsArray = !string.IsNullOrEmpty(arrayPart),
                ArrayBounds = string.IsNullOrEmpty(arrayPart) ? null : arrayPart.Trim('(', ')'),
                IsStatic = isStatic,
                LineNumber = token.LineNumber
            };
        }
    }

    private static VbaProcedure ParseProcedureOpen(Token token)
    {
        var g = token.Groups!;
        var kindStr = g[3].Value.Trim();
        var kind = kindStr.ToLowerInvariant() switch
        {
            "sub" => ProcedureKind.Sub,
            "function" => ProcedureKind.Function,
            var s when s.StartsWith("property") && s.EndsWith("get") => ProcedureKind.PropertyGet,
            var s when s.StartsWith("property") && s.EndsWith("let") => ProcedureKind.PropertyLet,
            var s when s.StartsWith("property") && s.EndsWith("set") => ProcedureKind.PropertySet,
            _ => ProcedureKind.Sub
        };

        var returnType = string.IsNullOrWhiteSpace(g[6].Value) ? null : g[6].Value.Trim();

        return new VbaProcedure
        {
            Scope = ParseAccessModifier(g[1].Value),
            IsStatic = !string.IsNullOrWhiteSpace(g[2].Value),
            Kind = kind,
            Name = g[4].Value.Trim(),
            Parameters = ParseParameters(g[5].Value),
            ReturnType = returnType,
            LineStart = token.LineNumber
        };
    }

    private static List<VbaParameter> ParseParameters(string paramString)
    {
        var result = new List<VbaParameter>();
        if (string.IsNullOrWhiteSpace(paramString)) return result;

        foreach (var part in SplitDeclarations(paramString))
        {
            var p = part.Trim();
            if (string.IsNullOrEmpty(p)) continue;

            var m = ReParam.Match(p);
            if (!m.Success) continue;

            var isOptional = m.Groups[1].Success;
            var isParamArray = m.Groups[2].Success;
            var passingStr = m.Groups[3].Value;
            var isByRef = !passingStr.Equals("ByVal", StringComparison.OrdinalIgnoreCase);
            var name = m.Groups[4].Value;
            var isArray = m.Groups[5].Success;
            var dataType = string.IsNullOrWhiteSpace(m.Groups[6].Value) ? "Variant" : m.Groups[6].Value.TrimEnd('(', ')');
            var defaultValue = m.Groups[7].Success ? m.Groups[7].Value.Trim() : null;

            result.Add(new VbaParameter
            {
                Name = name,
                DataType = dataType,
                IsByRef = isByRef,
                IsOptional = isOptional,
                IsParamArray = isParamArray,
                DefaultValue = defaultValue
            });
        }

        return result;
    }

    private static AccessModifier ParseAccessModifier(string value) =>
        value.Trim().ToLowerInvariant() switch
        {
            "public" => AccessModifier.Public,
            "private" => AccessModifier.Private,
            "friend" => AccessModifier.Friend,
            _ => AccessModifier.Default
        };

    /// <summary>
    /// Splits a comma-separated declaration list, respecting quoted strings.
    /// </summary>
    private static List<string> SplitDeclarations(string input)
    {
        var result = new List<string>();
        var current = new StringBuilder();
        var inString = false;

        foreach (var ch in input)
        {
            if (ch == '"') inString = !inString;
            if (ch == ',' && !inString)
            {
                result.Add(current.ToString());
                current.Clear();
            }
            else
            {
                current.Append(ch);
            }
        }

        if (current.Length > 0)
            result.Add(current.ToString());

        return result;
    }

    /// <summary>
    /// Extracts any trailing inline comment from a code line and appends it to the list.
    /// </summary>
    private static void AppendInlineComments(string line, List<string> target)
    {
        var idx = IndexOfInlineComment(line);
        if (idx >= 0)
        {
            var comment = line[(idx + 1)..].Trim();
            if (!string.IsNullOrEmpty(comment))
                target.Add(comment);
        }
    }

    private static int IndexOfInlineComment(string line)
    {
        var inString = false;
        for (var i = 0; i < line.Length; i++)
        {
            if (line[i] == '"') inString = !inString;
            if (!inString && line[i] == '\'') return i;
        }
        return -1;
    }
}
