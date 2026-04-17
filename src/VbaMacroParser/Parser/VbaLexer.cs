using System.Text.RegularExpressions;

namespace VbaMacroParser.Parser;

public enum TokenKind
{
    Attribute,
    Option,
    Const,
    Variable,
    ProcedureOpen,
    ProcedureClose,
    Comment,
    Blank,
    Other
}

public sealed class Token
{
    public TokenKind Kind { get; init; }
    public string Raw { get; init; } = string.Empty;
    public int LineNumber { get; init; }
    public GroupCollection? Groups { get; init; }
}

/// <summary>
/// Single-pass line-by-line tokeniser. Each line produces exactly one Token.
/// </summary>
public static class VbaLexer
{
    private static readonly Regex ReAttribute = new(
        @"^Attribute\s+VB_Name\s*=\s*""([^""]+)""",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);

    private static readonly Regex ReOption = new(
        @"^Option\s+(.+)",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);

    private static readonly Regex ReConst = new(
        @"^(?:(Public|Private|Friend)\s+)?Const\s+(\w+)(?:\s+As\s+(\w+))?\s*=\s*(.+)",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);

    private static readonly Regex ReVariable = new(
        @"^(Dim|Public|Private|Friend|Global|Static)\s+(.+)",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);

    // Allows one level of nested () for ParamArray parameters like items()
    private static readonly Regex ReProcedureOpen = new(
        @"^(?:(Public|Private|Friend)\s+)?(?:(Static)\s+)?(Sub|Function|Property\s+(?:Get|Let|Set))\s+(\w+)\s*\(([^()]*(?:\([^()]*\)[^()]*)*)\)(?:\s+As\s+(\w+))?",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);

    private static readonly Regex ReProcedureClose = new(
        @"^End\s+(Sub|Function|Property)",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);

    private static readonly Regex ReComment = new(
        @"^\s*(?:'|Rem\s)(.*)",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);

    public static IEnumerable<Token> Tokenise(IEnumerable<string> lines)
    {
        int lineNumber = 0;
        foreach (var raw in JoinContinuationLines(lines))
        {
            lineNumber++;
            var trimmed = raw.Trim();

            if (string.IsNullOrWhiteSpace(trimmed))
            {
                yield return new Token { Kind = TokenKind.Blank, Raw = raw, LineNumber = lineNumber };
                continue;
            }

            Match m;

            m = ReComment.Match(trimmed);
            if (m.Success)
            {
                yield return new Token { Kind = TokenKind.Comment, Raw = raw, LineNumber = lineNumber, Groups = m.Groups };
                continue;
            }

            m = ReAttribute.Match(trimmed);
            if (m.Success)
            {
                yield return new Token { Kind = TokenKind.Attribute, Raw = raw, LineNumber = lineNumber, Groups = m.Groups };
                continue;
            }

            m = ReOption.Match(trimmed);
            if (m.Success)
            {
                yield return new Token { Kind = TokenKind.Option, Raw = raw, LineNumber = lineNumber, Groups = m.Groups };
                continue;
            }

            m = ReConst.Match(trimmed);
            if (m.Success)
            {
                yield return new Token { Kind = TokenKind.Const, Raw = raw, LineNumber = lineNumber, Groups = m.Groups };
                continue;
            }

            m = ReProcedureOpen.Match(trimmed);
            if (m.Success)
            {
                yield return new Token { Kind = TokenKind.ProcedureOpen, Raw = raw, LineNumber = lineNumber, Groups = m.Groups };
                continue;
            }

            m = ReProcedureClose.Match(trimmed);
            if (m.Success)
            {
                yield return new Token { Kind = TokenKind.ProcedureClose, Raw = raw, LineNumber = lineNumber, Groups = m.Groups };
                continue;
            }

            m = ReVariable.Match(trimmed);
            if (m.Success)
            {
                yield return new Token { Kind = TokenKind.Variable, Raw = raw, LineNumber = lineNumber, Groups = m.Groups };
                continue;
            }

            yield return new Token { Kind = TokenKind.Other, Raw = raw, LineNumber = lineNumber };
        }
    }

    /// <summary>
    /// Joins VBA line-continuation sequences (lines ending with " _") into a single logical line.
    /// The line number of the first physical line is preserved.
    /// </summary>
    private static IEnumerable<string> JoinContinuationLines(IEnumerable<string> lines)
    {
        var buffer = new System.Text.StringBuilder();
        foreach (var line in lines)
        {
            var trimmedEnd = line.TrimEnd();
            if (trimmedEnd.EndsWith(" _", StringComparison.Ordinal))
            {
                // Strip the trailing " _" and a leading space on the continuation part
                buffer.Append(trimmedEnd[..^2].TrimEnd());
                buffer.Append(' ');
            }
            else
            {
                buffer.Append(line);
                yield return buffer.ToString();
                buffer.Clear();
            }
        }

        if (buffer.Length > 0)
            yield return buffer.ToString();
    }
}
