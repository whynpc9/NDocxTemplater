using NDocxTemplater;

var warnings = new List<RenderWarning>();
var options = new RenderOptions
{
    MissingValueBehavior = MissingValueBehavior.KeepTag,
    Culture = System.Globalization.CultureInfo.GetCultureInfo("de-DE"),
    WarningHandler = warnings.Add
};

var engine = new DocxTemplateEngine();
var output = engine.Render(File.ReadAllBytes("template.docx"), File.ReadAllText("data.json"), options);

File.WriteAllBytes("output.docx", output);
File.WriteAllLines("warnings.txt", warnings.Select(warning => $"{warning.Code}: {warning.Expression}"));
