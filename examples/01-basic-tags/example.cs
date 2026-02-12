using NDocxTemplater;

var engine = new DocxTemplateEngine();
var templateBytes = File.ReadAllBytes("template.docx");
var json = File.ReadAllText("data.json");
var output = engine.Render(templateBytes, json);
File.WriteAllBytes("output.docx", output);
