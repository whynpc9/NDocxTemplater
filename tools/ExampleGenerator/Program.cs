using System;
using System.Collections.Generic;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using NDocxTemplater;

var repoRoot = FindRepoRoot(AppContext.BaseDirectory);
var examplesRoot = Path.Combine(repoRoot, "examples");
Directory.CreateDirectory(examplesRoot);
const string TinyPngDataUri = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO8B9pYAAAAASUVORK5CYII=";

var engine = new DocxTemplateEngine();

GenerateExample(
    examplesRoot,
    "01-basic-tags",
    """
    {
      "patient": {
        "name": "Alice",
        "age": 34
      },
      "report": {
        "id": "RPT-001"
      }
    }
    """,
    """
    using NDocxTemplater;

    var engine = new DocxTemplateEngine();
    var templateBytes = File.ReadAllBytes("template.docx");
    var json = File.ReadAllText("data.json");
    var output = engine.Render(templateBytes, json);
    File.WriteAllBytes("output.docx", output);
    """,
    Paragraph("Patient: {patient.name}"),
    Paragraph("Age: {patient.age}"),
    Paragraph("Report ID: {report.id}"));

GenerateExample(
    examplesRoot,
    "02-condition",
    """
    {
      "patient": {
        "name": "Bob",
        "isVip": true
      }
    }
    """,
    """
    using NDocxTemplater;

    var engine = new DocxTemplateEngine();
    var output = engine.Render(File.ReadAllBytes("template.docx"), File.ReadAllText("data.json"));
    File.WriteAllBytes("output.docx", output);
    """,
    Paragraph("Patient: {patient.name}"),
    Paragraph("{?patient.isVip}"),
    Paragraph("VIP customer"),
    Paragraph("{/?patient.isVip}"),
    Paragraph("End."));

GenerateExample(
    examplesRoot,
    "03-loop",
    """
    {
      "orders": [
        { "id": "ORD-1", "amount": 120.5 },
        { "id": "ORD-2", "amount": 80 },
        { "id": "ORD-3", "amount": 66.2 }
      ]
    }
    """,
    """
    using NDocxTemplater;

    var engine = new DocxTemplateEngine();
    var output = engine.Render(File.ReadAllBytes("template.docx"), File.ReadAllText("data.json"));
    File.WriteAllBytes("output.docx", output);
    """,
    Paragraph("Orders"),
    Paragraph("{#orders}"),
    Paragraph("- {id}: {amount}"),
    Paragraph("{/orders}"));

GenerateExample(
    examplesRoot,
    "04-table-loop",
    """
    {
      "invoice": {
        "no": "INV-2026-001",
        "lines": [
          { "name": "Apple", "qty": 2, "price": 3.5 },
          { "name": "Banana", "qty": 3, "price": 2.2 },
          { "name": "Orange", "qty": 5, "price": 1.8 }
        ]
      }
    }
    """,
    """
    using NDocxTemplater;

    var engine = new DocxTemplateEngine();
    var output = engine.Render(File.ReadAllBytes("template.docx"), File.ReadAllText("data.json"));
    File.WriteAllBytes("output.docx", output);
    """,
    Paragraph("Invoice: {invoice.no}"),
    CreateTableTemplate());

GenerateExample(
    examplesRoot,
    "05-extensions",
    """
    {
      "reportDate": "2026-02-10T16:45:30Z",
      "orders": [
        { "id": "ORD-001", "amount": 12.5 },
        { "id": "ORD-002", "amount": 100 },
        { "id": "ORD-003", "amount": 66.2 }
      ]
    }
    """,
    """
    using NDocxTemplater;

    var engine = new DocxTemplateEngine();
    var output = engine.Render(File.ReadAllBytes("template.docx"), File.ReadAllText("data.json"));
    File.WriteAllBytes("output.docx", output);
    """,
    Paragraph("Orders count: {orders|count}"),
    Paragraph("Report date: {reportDate|format:date:yyyy-MM-dd}"),
    Paragraph("{#orders|sort:amount:desc|take:2}"),
    Paragraph("{id} -> {amount|format:number:0.00}"),
    Paragraph("{/orders|sort:amount:desc|take:2}"));

GenerateExample(
    examplesRoot,
    "06-images",
    """
    {
      "logo": {
        "src": "__TINY_PNG__",
        "width": 48,
        "height": 24
      },
      "gallery": [
        { "photo": { "src": "__TINY_PNG__", "width": 24, "height": 24 } },
        { "photo": { "src": "__TINY_PNG__", "width": 32, "height": 20 } }
      ]
    }
    """.Replace("__TINY_PNG__", TinyPngDataUri),
    """
    using NDocxTemplater;

    var engine = new DocxTemplateEngine();
    var output = engine.Render(File.ReadAllBytes("template.docx"), File.ReadAllText("data.json"));
    File.WriteAllBytes("output.docx", output);
    """,
    Paragraph("Inline logo"),
    Paragraph("{%logo}"),
    Paragraph("Gallery"),
    Paragraph("{#gallery}"),
    Paragraph("{%%photo}"),
    Paragraph("{/gallery}"));

Console.WriteLine($"Generated examples in: {examplesRoot}");

return;

static string FindRepoRoot(string startPath)
{
    var current = new DirectoryInfo(startPath);
    while (current != null)
    {
        var solutionPath = Path.Combine(current.FullName, "NDocxTemplater.sln");
        if (File.Exists(solutionPath))
        {
            return current.FullName;
        }

        current = current.Parent;
    }

    throw new InvalidOperationException("Cannot locate repository root from runtime path.");
}

static void GenerateExample(
    string examplesRoot,
    string name,
    string json,
    string sampleCode,
    params OpenXmlElement[] bodyElements)
{
    var dir = Path.Combine(examplesRoot, name);
    Directory.CreateDirectory(dir);

    var templatePath = Path.Combine(dir, "template.docx");
    var dataPath = Path.Combine(dir, "data.json");
    var outputPath = Path.Combine(dir, "output.docx");
    var codePath = Path.Combine(dir, "example.cs");

    File.WriteAllText(dataPath, json.Trim() + Environment.NewLine);
    File.WriteAllText(codePath, sampleCode.Trim() + Environment.NewLine);

    var templateBytes = CreateTemplate(bodyElements);
    File.WriteAllBytes(templatePath, templateBytes);

    var engine = new DocxTemplateEngine();
    var outputBytes = engine.Render(templateBytes, File.ReadAllText(dataPath));
    File.WriteAllBytes(outputPath, outputBytes);
}

static byte[] CreateTemplate(params OpenXmlElement[] bodyElements)
{
    using var stream = new MemoryStream();
    using (var document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document, true))
    {
        var mainPart = document.AddMainDocumentPart();
        var body = new Body();

        foreach (var element in bodyElements)
        {
            body.Append(element);
        }

        mainPart.Document = new Document(body);
        mainPart.Document.Save();
    }

    return stream.ToArray();
}

static Paragraph Paragraph(string text)
{
    return new Paragraph(new Run(new Text(text)));
}

static Table CreateTableTemplate()
{
    var table = new Table();
    table.Append(TableRow(Cell("Item"), Cell("Qty"), Cell("Unit Price")));
    table.Append(TableRow(Cell("{#invoice.lines}"), Cell(string.Empty), Cell(string.Empty)));
    table.Append(TableRow(Cell("{name}"), Cell("{qty}"), Cell("{price|format:number:0.00}")));
    table.Append(TableRow(Cell("{/invoice.lines}"), Cell(string.Empty), Cell(string.Empty)));
    return table;
}

static TableRow TableRow(params TableCell[] cells)
{
    return new TableRow(cells);
}

static TableCell Cell(string text)
{
    return new TableCell(new Paragraph(new Run(new Text(text))));
}
