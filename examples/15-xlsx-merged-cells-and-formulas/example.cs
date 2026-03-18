using NDocxTemplater;

var engine = new XlsxTemplateEngine();
var templateBytes = File.ReadAllBytes("template.xlsx");
var json = File.ReadAllText("data.json");
var output = engine.Render(templateBytes, json);
File.WriteAllBytes("output.xlsx", output);
