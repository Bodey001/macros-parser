using VbaMacroParser.Models;

namespace VbaMacroParser.Parser;

public interface IVbaParser
{
    VbaParseResult Parse(string sourceFilePath, IEnumerable<string> lines);
}
