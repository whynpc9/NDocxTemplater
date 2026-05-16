using NDocxTemplater;

var engine = new PptxTemplateEngine();
var templateBytes = File.ReadAllBytes("template.pptx");
var json = File.ReadAllText("data.json");
var output = engine.Render(templateBytes, json);
File.WriteAllBytes("output.pptx", output);
