using NDocxTemplater;

var engine = new DocxTemplateEngine();
var output = engine.Render(File.ReadAllBytes("template.docx"), File.ReadAllText("data.json"));
File.WriteAllBytes("output.docx", output);
