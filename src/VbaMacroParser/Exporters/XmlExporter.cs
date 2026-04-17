using System.Text;
using System.Xml;
using VbaMacroParser.Models;

namespace VbaMacroParser.Exporters;

public sealed class XmlExporter : IExporter
{
    public string FileExtension => ".xml";

    public string Export(VbaParseResult result)
    {
        var sb = new StringBuilder();
        var settings = new XmlWriterSettings
        {
            Indent = true,
            IndentChars = "  ",
            Encoding = new UTF8Encoding(encoderShouldEmitUTF8Identifier: false),
            OmitXmlDeclaration = false
        };

        using var writer = XmlWriter.Create(sb, settings);
        writer.WriteStartDocument();
        writer.WriteStartElement("VbaParseResult");
        writer.WriteAttributeString("parserVersion", result.ParserVersion);
        writer.WriteAttributeString("parsedAt", result.ParsedAt.ToString("O"));
        writer.WriteAttributeString("sourceFile", result.SourceFile);

        writer.WriteStartElement("Modules");
        foreach (var module in result.Modules)
            WriteModule(writer, module);
        writer.WriteEndElement(); // Modules

        writer.WriteEndElement(); // VbaParseResult
        writer.WriteEndDocument();

        writer.Flush();
        return sb.ToString();
    }

    private static void WriteModule(XmlWriter w, VbaModule module)
    {
        w.WriteStartElement("Module");
        w.WriteAttributeString("name", module.Name);
        w.WriteAttributeString("type", module.Type.ToString());

        if (module.Options.Count > 0)
        {
            w.WriteStartElement("Options");
            foreach (var opt in module.Options)
                w.WriteElementString("Option", opt);
            w.WriteEndElement();
        }

        if (module.ModuleComments.Count > 0)
        {
            w.WriteStartElement("ModuleComments");
            foreach (var c in module.ModuleComments)
                w.WriteElementString("Comment", c);
            w.WriteEndElement();
        }

        if (module.Constants.Count > 0)
        {
            w.WriteStartElement("Constants");
            foreach (var c in module.Constants)
            {
                w.WriteStartElement("Constant");
                w.WriteAttributeString("name", c.Name);
                w.WriteAttributeString("scope", c.Scope.ToString());
                w.WriteAttributeString("dataType", c.DataType);
                w.WriteAttributeString("lineNumber", c.LineNumber.ToString());
                w.WriteString(c.Value);
                w.WriteEndElement();
            }
            w.WriteEndElement();
        }

        if (module.Variables.Count > 0)
        {
            w.WriteStartElement("Variables");
            foreach (var v in module.Variables)
            {
                w.WriteStartElement("Variable");
                w.WriteAttributeString("name", v.Name);
                w.WriteAttributeString("scope", v.Scope.ToString());
                w.WriteAttributeString("dataType", v.DataType);
                w.WriteAttributeString("isArray", v.IsArray.ToString().ToLower());
                if (v.ArrayBounds != null)
                    w.WriteAttributeString("arrayBounds", v.ArrayBounds);
                w.WriteAttributeString("isStatic", v.IsStatic.ToString().ToLower());
                w.WriteAttributeString("lineNumber", v.LineNumber.ToString());
                w.WriteEndElement();
            }
            w.WriteEndElement();
        }

        if (module.Procedures.Count > 0)
        {
            w.WriteStartElement("Procedures");
            foreach (var proc in module.Procedures)
                WriteProcedure(w, proc);
            w.WriteEndElement();
        }

        w.WriteEndElement(); // Module
    }

    private static void WriteProcedure(XmlWriter w, VbaProcedure proc)
    {
        w.WriteStartElement("Procedure");
        w.WriteAttributeString("name", proc.Name);
        w.WriteAttributeString("kind", proc.Kind.ToString());
        w.WriteAttributeString("scope", proc.Scope.ToString());
        w.WriteAttributeString("isStatic", proc.IsStatic.ToString().ToLower());
        w.WriteAttributeString("lineStart", proc.LineStart.ToString());
        w.WriteAttributeString("lineEnd", proc.LineEnd.ToString());
        if (proc.ReturnType != null)
            w.WriteAttributeString("returnType", proc.ReturnType);

        if (proc.Parameters.Count > 0)
        {
            w.WriteStartElement("Parameters");
            foreach (var p in proc.Parameters)
            {
                w.WriteStartElement("Parameter");
                w.WriteAttributeString("name", p.Name);
                w.WriteAttributeString("dataType", p.DataType);
                w.WriteAttributeString("isByRef", p.IsByRef.ToString().ToLower());
                w.WriteAttributeString("isOptional", p.IsOptional.ToString().ToLower());
                w.WriteAttributeString("isParamArray", p.IsParamArray.ToString().ToLower());
                if (p.DefaultValue != null)
                    w.WriteAttributeString("defaultValue", p.DefaultValue);
                w.WriteEndElement();
            }
            w.WriteEndElement();
        }

        if (proc.Comments.Count > 0)
        {
            w.WriteStartElement("Comments");
            foreach (var c in proc.Comments)
                w.WriteElementString("Comment", c);
            w.WriteEndElement();
        }

        if (!string.IsNullOrWhiteSpace(proc.Body))
            w.WriteElementString("Body", proc.Body);

        w.WriteEndElement(); // Procedure
    }
}
