using VbaMacroParser.IO;
using VbaMacroParser.Parser;

System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

var banner = """
  ╔══════════════════════════════════════╗
  ║     VBA Macro Parser  v1.0.0         ║
  ╚══════════════════════════════════════╝
""";

Console.WriteLine(banner);

// ---------------------------------------------------------------------------
// Resolve input files
// ---------------------------------------------------------------------------
string[] inputFiles;

if (args.Length >= 1)
{
    inputFiles = args;
}
else
{
    // Try to auto-detect files in the input/ folder next to the binary.
    var inputDir = ResolveInputDir();
    if (!Directory.Exists(inputDir))
    {
        Console.Error.WriteLine($"[ERROR] input/ folder not found at: {inputDir}");
        Console.Error.WriteLine("Usage: VbaMacroParser <path-to-vba-file>");
        return 2;
    }

    inputFiles = Directory.GetFiles(inputDir)
        .Where(f => !Path.GetFileName(f).StartsWith("."))
        .ToArray();
        
    if (inputFiles.Length == 0)
    {
        Console.Error.WriteLine("[ERROR] No files found in the input/ folder.");
        Console.Error.WriteLine("        Drop a VBA source file there, or pass a path as an argument.");
        return 2;
    }
}

var parser = new VbaParser();
var outputsRoot = ResolveOutputsDir();
var manager = new OutputManager(outputsRoot);

foreach (var inputFile in inputFiles)
{
    var absolutePath = Path.GetFullPath(inputFile);
    Console.WriteLine($"\nProcessing: {absolutePath}");

    try
    {
        var lines = FileReader.ReadLines(absolutePath);
        Console.WriteLine($"  Lines  : {lines.Length}");
        
        var result = parser.Parse(absolutePath, lines);

        var totalProcs = result.Modules.Sum(m => m.Procedures.Count);
        var totalVars  = result.Modules.Sum(m => m.Variables.Count);
        var totalConst = result.Modules.Sum(m => m.Constants.Count);

        Console.WriteLine($"  Modules: {result.Modules.Count}");
        Console.WriteLine($"  Procs  : {totalProcs}  Variables: {totalVars}  Constants: {totalConst}");

        var outputDir = manager.Run(result);
        
        Console.WriteLine($"  Output : {outputDir}");
        Console.WriteLine("  Files  : .json  .xml  .txt  .csv");
    }
    catch (UnsupportedFormatException ex)
    {
        Console.Error.WriteLine($"  [ERROR] {ex.Message}");
    }
    catch (FileNotFoundException ex)
    {
        Console.Error.WriteLine($"  [ERROR] {ex.Message}");
    }
    catch (Exception ex)
    {
        Console.Error.WriteLine($"  [ERROR] Failed to process {Path.GetFileName(absolutePath)}: {ex.Message}");
    }
}

Console.WriteLine("\n  Done.");
return 0;

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------
static string ResolveInputDir()
{
    // Walk up from the executable to find the repository root that has input/.
    var dir = AppContext.BaseDirectory;
    for (var i = 0; i < 8; i++)
    {
        var candidate = Path.Combine(dir, "input");
        if (Directory.Exists(candidate)) return candidate;
        var parent = Directory.GetParent(dir);
        if (parent is null) break;
        dir = parent.FullName;
    }
    return Path.Combine(AppContext.BaseDirectory, "input");
}

static string ResolveOutputsDir()
{
    var dir = AppContext.BaseDirectory;
    for (var i = 0; i < 8; i++)
    {
        var candidate = Path.Combine(dir, "output");
        if (Directory.Exists(candidate)) return candidate;
        // If input/ exists at this level, output/ should live here too.
        if (Directory.Exists(Path.Combine(dir, "input")))
            return Path.Combine(dir, "output");
        var parent = Directory.GetParent(dir);
        if (parent is null) break;
        dir = parent.FullName;
    }
    return Path.Combine(AppContext.BaseDirectory, "output");
}
