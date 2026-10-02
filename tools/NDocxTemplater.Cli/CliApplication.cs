using System.IO.Compression;
using System.Text.RegularExpressions;
using System.Xml.Linq;

namespace NDocxTemplater.Cli;

public static class CliApplication
{
    private static readonly Regex TagRegex = new Regex(@"\{([^{}]+)\}", RegexOptions.Compiled);

    public static int Run(string[] args, TextWriter output, TextWriter error)
    {
        try
        {
            if (args.Length == 0 || IsHelp(args[0]))
            {
                WriteHelp(output);
                return 0;
            }

            var command = args[0].Trim().ToLowerInvariant();
            var options = ParseOptions(args.Skip(1).ToArray());

            switch (command)
            {
                case "render":
                    return Render(options, output, error);
                case "validate":
                    return Validate(options, output, error);
                case "inspect-tags":
                    return InspectTags(options, output, error);
                default:
                    error.WriteLine("Unknown command: " + args[0]);
                    WriteHelp(error);
                    return 2;
            }
        }
        catch (Exception ex)
        {
            error.WriteLine(ex.Message);
            return 1;
        }
    }

    private static int Render(IReadOnlyDictionary<string, string> options, TextWriter output, TextWriter error)
    {
        var templatePath = RequireOption(options, "template");
        var dataPath = RequireOption(options, "data");
        var outputPath = RequireOption(options, "output");
        var format = GetFormat(options, templatePath);
        var renderOptions = BuildRenderOptions(options);

        var rendered = RenderBytes(format, File.ReadAllBytes(templatePath), File.ReadAllText(dataPath), renderOptions);
        Directory.CreateDirectory(Path.GetDirectoryName(Path.GetFullPath(outputPath)) ?? ".");
        File.WriteAllBytes(outputPath, rendered);

        output.WriteLine("Rendered " + format.ToUpperInvariant() + ": " + outputPath);
        return 0;
    }

    private static int Validate(IReadOnlyDictionary<string, string> options, TextWriter output, TextWriter error)
    {
        var templatePath = RequireOption(options, "template");
        var dataPath = RequireOption(options, "data");
        var format = GetFormat(options, templatePath);
        var warnings = new List<RenderWarning>();
        var renderOptions = BuildRenderOptions(options);
        renderOptions.WarningHandler = warning =>
        {
            warnings.Add(warning);
            if (options.TryGetValue("warnings", out var warningsPath))
            {
                File.AppendAllText(warningsPath, warning.Code + ": " + warning.Expression + Environment.NewLine);
            }
        };

        if (!options.ContainsKey("missing"))
        {
            renderOptions.MissingValueBehavior = MissingValueBehavior.Throw;
        }

        _ = RenderBytes(format, File.ReadAllBytes(templatePath), File.ReadAllText(dataPath), renderOptions);

        output.WriteLine("Template validation OK.");
        if (warnings.Count > 0)
        {
            output.WriteLine("Warnings: " + warnings.Count.ToString(System.Globalization.CultureInfo.InvariantCulture));
        }

        return 0;
    }

    private static int InspectTags(IReadOnlyDictionary<string, string> options, TextWriter output, TextWriter error)
    {
        var templatePath = RequireOption(options, "template");
        var tags = InspectTemplateTags(templatePath).ToArray();
        foreach (var tag in tags)
        {
            output.WriteLine(tag);
        }

        return 0;
    }

    public static IReadOnlyList<string> InspectTemplateTags(string templatePath)
    {
        using (var stream = File.OpenRead(templatePath))
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: false))
        {
            return archive.Entries
                .Where(static entry => entry.FullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase))
                .SelectMany(ReadTags)
                .Select(static tag => tag.Trim())
                .Where(static tag => tag.Length > 0)
                .Distinct(StringComparer.Ordinal)
                .OrderBy(static tag => tag, StringComparer.Ordinal)
                .ToArray();
        }
    }

    private static IEnumerable<string> ReadTags(ZipArchiveEntry entry)
    {
        using (var reader = new StreamReader(entry.Open()))
        {
            var xml = XDocument.Load(reader);
            XNamespace word = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
            XNamespace drawing = "http://schemas.openxmlformats.org/drawingml/2006/main";
            XNamespace spreadsheet = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
            var containers = xml.Descendants().Where(element =>
                element.Name == word + "p" || element.Name == drawing + "p"
                || element.Name == spreadsheet + "si" || element.Name == spreadsheet + "is"
                || element.Name == spreadsheet + "f");
            foreach (var container in containers)
            {
                var text = container.Name == spreadsheet + "f"
                    ? container.Value
                    : string.Concat(container.Descendants().Where(element =>
                        (element.Name == word + "t" || element.Name == drawing + "t" || element.Name == spreadsheet + "t")
                        && element.Ancestors().FirstOrDefault(ancestor =>
                            ancestor.Name == word + "p" || ancestor.Name == drawing + "p"
                            || ancestor.Name == spreadsheet + "si" || ancestor.Name == spreadsheet + "is") == container)
                        .Select(static element => element.Value));
                foreach (Match match in TagRegex.Matches(text))
                {
                    yield return match.Groups[1].Value;
                }
            }
        }
    }

    private static byte[] RenderBytes(string format, byte[] templateBytes, string jsonData, RenderOptions options)
    {
        switch (format)
        {
            case "docx":
                return new DocxTemplateEngine().Render(templateBytes, jsonData, options);
            case "xlsx":
                return new XlsxTemplateEngine().Render(templateBytes, jsonData, options);
            case "pptx":
                return new PptxTemplateEngine().Render(templateBytes, jsonData, options);
            default:
                throw new InvalidOperationException("Unsupported format: " + format);
        }
    }

    private static RenderOptions BuildRenderOptions(IReadOnlyDictionary<string, string> options)
    {
        var renderOptions = new RenderOptions();
        if (options.TryGetValue("base-dir", out var baseDirectory))
        {
            renderOptions.BaseDirectory = baseDirectory;
        }

        if (options.TryGetValue("missing", out var missingValue))
        {
            renderOptions.MissingValueBehavior = ParseMissingValueBehavior(missingValue);
        }

        if (options.TryGetValue("culture", out var culture))
        {
            renderOptions.Culture = System.Globalization.CultureInfo.GetCultureInfo(culture);
        }

        return renderOptions;
    }

    private static MissingValueBehavior ParseMissingValueBehavior(string value)
    {
        switch (value.Trim().ToLowerInvariant())
        {
            case "empty":
                return MissingValueBehavior.Empty;
            case "keep":
            case "keep-tag":
            case "keeptag":
                return MissingValueBehavior.KeepTag;
            case "throw":
            case "strict":
                return MissingValueBehavior.Throw;
            default:
                throw new InvalidOperationException("Unsupported missing value behavior: " + value);
        }
    }

    private static string GetFormat(IReadOnlyDictionary<string, string> options, string templatePath)
    {
        if (options.TryGetValue("format", out var format))
        {
            return NormalizeFormat(format);
        }

        return NormalizeFormat(Path.GetExtension(templatePath).TrimStart('.'));
    }

    private static string NormalizeFormat(string format)
    {
        var normalized = format.Trim().TrimStart('.').ToLowerInvariant();
        if (normalized == "docx" || normalized == "xlsx" || normalized == "pptx")
        {
            return normalized;
        }

        throw new InvalidOperationException("Format must be docx, xlsx, or pptx.");
    }

    private static string RequireOption(IReadOnlyDictionary<string, string> options, string name)
    {
        if (options.TryGetValue(name, out var value) && !string.IsNullOrWhiteSpace(value))
        {
            return value;
        }

        throw new InvalidOperationException("Missing required option --" + name + ".");
    }

    private static Dictionary<string, string> ParseOptions(IReadOnlyList<string> args)
    {
        var options = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        for (var index = 0; index < args.Count; index++)
        {
            var item = args[index];
            if (!item.StartsWith("--", StringComparison.Ordinal))
            {
                throw new InvalidOperationException("Unexpected argument: " + item);
            }

            var name = item.Substring(2);
            if (name.Length == 0)
            {
                throw new InvalidOperationException("Invalid option name.");
            }

            if (index + 1 >= args.Count || args[index + 1].StartsWith("--", StringComparison.Ordinal))
            {
                options[name] = "true";
                continue;
            }

            options[name] = args[++index];
        }

        return options;
    }

    private static bool IsHelp(string value)
    {
        return string.Equals(value, "-h", StringComparison.Ordinal)
            || string.Equals(value, "--help", StringComparison.OrdinalIgnoreCase)
            || string.Equals(value, "help", StringComparison.OrdinalIgnoreCase);
    }

    private static void WriteHelp(TextWriter writer)
    {
        writer.WriteLine("NDocxTemplater CLI");
        writer.WriteLine();
        writer.WriteLine("Commands:");
        writer.WriteLine("  render --template TEMPLATE --data DATA --output OUTPUT [--format docx|xlsx|pptx]");
        writer.WriteLine("  validate --template TEMPLATE --data DATA [--format docx|xlsx|pptx]");
        writer.WriteLine("  inspect-tags --template TEMPLATE");
        writer.WriteLine();
        writer.WriteLine("Options:");
        writer.WriteLine("  --missing empty|keep|throw");
        writer.WriteLine("  --base-dir PATH");
        writer.WriteLine("  --culture CULTURE");
    }
}
