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
var realChartAssetPath = Path.Combine(repoRoot, "tests", "NDocxTemplater.Tests", "Assets", "real-chart.png");

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

GenerateExample(
    examplesRoot,
    "07-table-date-format-split-runs",
    """
    {
      "rows": [
        { "name": "Row A", "createdAt": "2026-02-24T10:11:12Z" },
        { "name": "Row B", "createdAt": "2026-03-01T01:02:03Z" }
      ]
    }
    """,
    """
    using NDocxTemplater;

    var engine = new DocxTemplateEngine();
    var output = engine.Render(File.ReadAllBytes("template.docx"), File.ReadAllText("data.json"));
    File.WriteAllBytes("output.docx", output);
    """,
    new Table(
        TableRow(Cell("Name"), Cell("Created")),
        TableRow(Cell("{#rows}"), Cell(string.Empty)),
        TableRow(
            Cell("{name}"),
            CellWithSplitRuns("{createdAt|for", "mat:date:yyyy-MM-", "dd}")),
        TableRow(Cell("{/rows}"), Cell(string.Empty))));

GenerateExample(
    examplesRoot,
    "08-inline-friendly-expressions",
    """
    {
      "financeMonthly": [
        { "month": "2025-03-01", "revenue": 90000 },
        { "month": "2025-01-01", "revenue": 70000 },
        { "month": "2025-07-01", "revenue": 85000 },
        { "month": "2025-05-01", "revenue": 100000 }
      ],
      "institutions": [
        { "name": "机构C", "revenue": 650000 },
        { "name": "机构A", "revenue": 1000000 },
        { "name": "机构Z", "revenue": 100000 }
      ]
    }
    """,
    """
    using NDocxTemplater;

    var engine = new DocxTemplateEngine();
    var output = engine.Render(File.ReadAllBytes("template.docx"), File.ReadAllText("data.json"));
    File.WriteAllBytes("output.docx", output);
    """,
    Paragraph(
        "统计数据包括了从{financeMonthly|sort:month:asc|first|get:month|format:date:yyyy年M月}到{financeMonthly|sort:month:asc|last|get:month|format:date:yyyy年M月}的财务数据，其中营收最高的是{financeMonthly|maxby:revenue|get:month|format:date:M月}，营收为{financeMonthly|maxby:revenue|get:revenue|format:number:#,##0}元"),
    Paragraph(
        "在这些机构的对比数据中，其中营收最高的为{institutions|maxby:revenue|get:name}，收入为{institutions|maxby:revenue|get:revenue|format:number:#,##0}元，营收最低的为{institutions|minby:revenue|get:name}，收入为{institutions|minby:revenue|get:revenue|format:number:#,##0}元"));

GenerateExample(
    examplesRoot,
    "09-inline-ranking-positions",
    """
    {
      "institutions": [
        { "name": "机构A", "revenue": 1000000 },
        { "name": "机构B", "revenue": 920000 },
        { "name": "机构C", "revenue": 880000 },
        { "name": "机构D", "revenue": 860000 },
        { "name": "机构E", "revenue": 840000 },
        { "name": "机构F", "revenue": 820000 },
        { "name": "机构G", "revenue": 800000 },
        { "name": "机构H", "revenue": 780000 },
        { "name": "机构I", "revenue": 760000 },
        { "name": "机构J", "revenue": 740000 },
        { "name": "机构K", "revenue": 100000 }
      ]
    }
    """,
    """
    using NDocxTemplater;

    var engine = new DocxTemplateEngine();
    var output = engine.Render(File.ReadAllBytes("template.docx"), File.ReadAllText("data.json"));
    File.WriteAllBytes("output.docx", output);
    """,
    Paragraph("前10名机构中，第3名为{institutions|sort:revenue:desc|take:10|nth:3|get:name}，收入为{institutions|sort:revenue:desc|take:10|nth:3|get:revenue|format:number:#,##0}元。"),
    Paragraph("前10名末位为{institutions|sort:revenue:desc|take:10|at:-1|get:name}（支持负数索引）。"));

GenerateExample(
    examplesRoot,
    "10-inline-conditions-and-rates",
    """
    {
      "flags": {
        "includeRates": true
      },
      "metrics": {
        "growthRate": 0.0123,
        "badDebtRate": 0.0045
      },
      "institutions": [
        { "name": "机构A" },
        { "name": "机构B" },
        { "name": "机构C" }
      ]
    }
    """,
    """
    using NDocxTemplater;

    var engine = new DocxTemplateEngine();
    var output = engine.Render(File.ReadAllBytes("template.docx"), File.ReadAllText("data.json"));
    File.WriteAllBytes("output.docx", output);
    """,
    Paragraph("本次样本共{institutions|count}家机构，状态：{flags.includeRates|if:包含比率指标:不包含比率指标}，环比增长率{metrics.growthRate|format:percent:0.00}，坏账率{metrics.badDebtRate|format:permille:0.00}。"),
    Paragraph("备用写法（number pattern）：{metrics.growthRate|format:number:0.00%} / {metrics.badDebtRate|format:number:0.00‰}"));

GenerateImagePathAndDataUriScalingExample(examplesRoot, realChartAssetPath);

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
    var originalCurrentDirectory = Environment.CurrentDirectory;
    Environment.CurrentDirectory = dir;
    byte[] outputBytes;
    try
    {
        outputBytes = engine.Render(templateBytes, File.ReadAllText(dataPath));
    }
    finally
    {
        Environment.CurrentDirectory = originalCurrentDirectory;
    }

    File.WriteAllBytes(outputPath, outputBytes);
}

static void GenerateImagePathAndDataUriScalingExample(string examplesRoot, string imageAssetPath)
{
    if (!File.Exists(imageAssetPath))
    {
        throw new FileNotFoundException("Missing real image asset for examples.", imageAssetPath);
    }

    var imageBytes = File.ReadAllBytes(imageAssetPath);
    var imageDataUri = "data:image/png;base64," + Convert.ToBase64String(imageBytes);

    var dir = Path.Combine(examplesRoot, "11-images-file-and-datauri-scaling");
    Directory.CreateDirectory(dir);

    var exampleImagePath = Path.Combine(dir, "chart.png");
    File.Copy(imageAssetPath, exampleImagePath, overwrite: true);

    GenerateExample(
        examplesRoot,
        "11-images-file-and-datauri-scaling",
        """
        {
          "fromPath": {
            "src": "chart.png",
            "maxWidth": 376,
            "preserveAspectRatio": true
          },
          "fromDataUri": {
            "src": "__REAL_CHART_DATA_URI__",
            "scale": 0.25,
            "preserveAspectRatio": true
          },
          "fitInBox": {
            "src": "chart.png",
            "width": 420,
            "height": 260,
            "preserveAspectRatio": true
          }
        }
        """.Replace("__REAL_CHART_DATA_URI__", imageDataUri),
        """
        using NDocxTemplater;

        var engine = new DocxTemplateEngine();
        var output = engine.Render(File.ReadAllBytes("template.docx"), File.ReadAllText("data.json"));
        File.WriteAllBytes("output.docx", output);
        """,
        Paragraph("图片（文件路径 + maxWidth 等比缩放）"),
        Paragraph("{%%fromPath}"),
        Paragraph("图片（data URI + scale 等比缩放）"),
        Paragraph("{%%fromDataUri}"),
        Paragraph("图片（在固定宽高盒子中等比缩放，不拉伸变形）"),
        Paragraph("{%%fitInBox}"));
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

static TableCell CellWithSplitRuns(params string[] pieces)
{
    var paragraph = new Paragraph();
    foreach (var piece in pieces)
    {
        paragraph.Append(new Run(new Text(piece)));
    }

    return new TableCell(paragraph);
}
