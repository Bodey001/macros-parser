using VbaMacroParser.Exporters;
using VbaMacroParser.Models;

namespace VbaMacroParser.IO;

public sealed class OutputManager
{
    private readonly IReadOnlyList<IExporter> _exporters;
    private readonly string _outputsRoot;

    public OutputManager(string outputsRoot, IReadOnlyList<IExporter>? exporters = null)
    {
        _outputsRoot = outputsRoot;
        _exporters = exporters ?? CreateDefaultExporters();
    }

    public string Run(VbaParseResult result)
    {
        var baseName = Path.GetFileNameWithoutExtension(result.SourceFile);
        var subDir = Path.Combine(_outputsRoot, $"{baseName}-parsed");
        Directory.CreateDirectory(subDir);

        foreach (var exporter in _exporters)
        {
            var outputPath = Path.Combine(subDir, baseName + exporter.FileExtension);
            var content = exporter.Export(result);
            File.WriteAllText(outputPath, content, System.Text.Encoding.UTF8);
        }

        return subDir;
    }

    private static IReadOnlyList<IExporter> CreateDefaultExporters() =>
    [
        new JsonExporter(),
        new XmlExporter(),
        new TextExporter(),
        new CsvExporter()
    ];
}
