using System;
using System.Collections.Generic;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using NDocxTemplater;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;
using S = DocumentFormat.OpenXml.Spreadsheet;

var repoRoot = FindRepoRoot(AppContext.BaseDirectory);
var examplesRoot = Path.Combine(repoRoot, "examples");
Directory.CreateDirectory(examplesRoot);
const string TinyPngDataUri = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO8B9pYAAAAASUVORK5CYII=";
var realChartAssetPath = Path.Combine(repoRoot, "tests", "NDocxTemplater.Tests", "Assets", "real-chart.png");

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

GenerateExample(
    examplesRoot,
    "12-barcodes",
    """
    {
      "barcodes": {
        "code128": "A20260303001",
        "ean13": "6901234567892"
      }
    }
    """,
    """
    using NDocxTemplater;

    var engine = new DocxTemplateEngine();
    var output = engine.Render(File.ReadAllBytes("template.docx"), File.ReadAllText("data.json"));
    File.WriteAllBytes("output.docx", output);
    """,
    Paragraph("Code128 条形码（由模板参数指定类型和尺寸）"),
    Paragraph("{%barcode:barcodes.code128;type=code128;width=360;height=96;margin=1}"),
    Paragraph("EAN13 条形码（居中）"),
    Paragraph("{%%barcode:barcodes.ean13;type=ean13;width=280;height=100;margin=2}"));

GenerateWorkbookExample(
    examplesRoot,
    "13-xlsx-row-loop",
    """
    {
      "report": {
        "title": "Sales Summary",
        "lines": [
          { "name": "Apple", "qty": 2, "amount": 12.5 },
          { "name": "Banana", "qty": 5, "amount": 66.2 },
          { "name": "Orange", "qty": 3, "amount": 100.0 }
        ]
      },
      "showFooter": true
    }
    """,
    """
    using NDocxTemplater;

    var engine = new XlsxTemplateEngine();
    var templateBytes = File.ReadAllBytes("template.xlsx");
    var json = File.ReadAllText("data.json");
    var output = engine.Render(templateBytes, json);
    File.WriteAllBytes("output.xlsx", output);
    """,
    WorkbookRowSpec.Create("Report", "{report.title}"),
    WorkbookRowSpec.Create("Item", "Qty", "Amount", styleIndex: 1U),
    WorkbookRowSpec.Create("{#report.lines|sort:amount:desc|take:2}", string.Empty, string.Empty),
    WorkbookRowSpec.Create("{name}", "{qty}", "{amount|format:number:0.00}", styleIndex: 1U),
    WorkbookRowSpec.Create("{/report.lines|sort:amount:desc|take:2}", string.Empty, string.Empty),
    WorkbookRowSpec.Create("{?showFooter}", string.Empty, string.Empty),
    WorkbookRowSpec.Create("Count", "{report.lines|count}", string.Empty, styleIndex: 1U),
    WorkbookRowSpec.Create("{/?showFooter}", string.Empty, string.Empty));

GenerateWorkbookMediaExample(examplesRoot, realChartAssetPath);
GenerateWorkbookMergeAndFormulaExample(examplesRoot);
GeneratePresentationExample(
    examplesRoot,
    "16-pptx-slide-loop-and-condition",
    """
    {
      "report": {
        "title": "Quarterly Review"
      },
      "users": [
        { "name": "Alice", "amount": 12.5 },
        { "name": "Bob", "amount": 99 }
      ],
      "showSummary": true,
      "summaryText": "All regions green"
    }
    """,
    """
    using NDocxTemplater;

    var engine = new PptxTemplateEngine();
    var templateBytes = File.ReadAllBytes("template.pptx");
    var json = File.ReadAllText("data.json");
    var output = engine.Render(templateBytes, json);
    File.WriteAllBytes("output.pptx", output);
    """,
    PresentationSlideSpec.Create("Title {report.title}"),
    PresentationSlideSpec.Create("{:users}", "User {name}", "Amount {amount|format:number:0.00}"),
    PresentationSlideSpec.Create("{:showSummary}", "Summary {summaryText}"));

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

static void GenerateWorkbookExample(
    string examplesRoot,
    string name,
    string json,
    string sampleCode,
    params WorkbookRowSpec[] rows)
{
    var dir = Path.Combine(examplesRoot, name);
    Directory.CreateDirectory(dir);

    var templatePath = Path.Combine(dir, "template.xlsx");
    var dataPath = Path.Combine(dir, "data.json");
    var outputPath = Path.Combine(dir, "output.xlsx");
    var codePath = Path.Combine(dir, "example.cs");

    File.WriteAllText(dataPath, json.Trim() + Environment.NewLine);
    File.WriteAllText(codePath, sampleCode.Trim() + Environment.NewLine);

    var templateBytes = CreateWorkbookTemplate(rows);
    File.WriteAllBytes(templatePath, templateBytes);

    var engine = new XlsxTemplateEngine();
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

static void GenerateWorkbookMediaExample(string examplesRoot, string imageAssetPath)
{
    if (!File.Exists(imageAssetPath))
    {
        throw new FileNotFoundException("Missing real image asset for workbook examples.", imageAssetPath);
    }

    var imageBytes = File.ReadAllBytes(imageAssetPath);
    var imageDataUri = "data:image/png;base64," + Convert.ToBase64String(imageBytes);
    var dir = Path.Combine(examplesRoot, "14-xlsx-media-placeholders");
    Directory.CreateDirectory(dir);

    var chartPath = Path.Combine(dir, "chart.png");
    File.Copy(imageAssetPath, chartPath, overwrite: true);

    GenerateWorkbookExample(
        examplesRoot,
        "14-xlsx-media-placeholders",
        """
        {
          "pathImage": {
            "src": "chart.png",
            "maxWidth": 320,
            "preserveAspectRatio": true
          },
          "dataUriImage": {
            "src": "__REAL_CHART_DATA_URI__",
            "scale": 0.2,
            "preserveAspectRatio": true
          },
          "barcodes": {
            "ean13": "5901234123457"
          }
        }
        """.Replace("__REAL_CHART_DATA_URI__", imageDataUri),
        """
        using NDocxTemplater;

        var engine = new XlsxTemplateEngine();
        var templateBytes = File.ReadAllBytes("template.xlsx");
        var json = File.ReadAllText("data.json");
        var output = engine.Render(templateBytes, json);
        File.WriteAllBytes("output.xlsx", output);
        """,
        WorkbookRowSpec.Create("Media", "Value"),
        WorkbookRowSpec.Create("From path", "{%pathImage}"),
        WorkbookRowSpec.Create("From data URI", "{%dataUriImage}"),
        WorkbookRowSpec.Create("Barcode", "{%barcode:barcodes.ean13;type=ean13;width=220;height=80}"));
}

static void GenerateWorkbookMergeAndFormulaExample(string examplesRoot)
{
    var dir = Path.Combine(examplesRoot, "15-xlsx-merged-cells-and-formulas");
    Directory.CreateDirectory(dir);

    var templatePath = Path.Combine(dir, "template.xlsx");
    var dataPath = Path.Combine(dir, "data.json");
    var outputPath = Path.Combine(dir, "output.xlsx");
    var codePath = Path.Combine(dir, "example.cs");

    var json = """
        {
          "lines": [
            { "name": "Alpha", "qty": 2, "price": 5 },
            { "name": "Beta", "qty": 3, "price": 7 }
          ]
        }
        """;

    var sampleCode = """
        using NDocxTemplater;

        var engine = new XlsxTemplateEngine();
        var templateBytes = File.ReadAllBytes("template.xlsx");
        var json = File.ReadAllText("data.json");
        var output = engine.Render(templateBytes, json);
        File.WriteAllBytes("output.xlsx", output);
        """;

    File.WriteAllText(dataPath, json.Trim() + Environment.NewLine);
    File.WriteAllText(codePath, sampleCode.Trim() + Environment.NewLine);

    var templateBytes = CreateAdvancedWorkbookTemplate(
        new[] { "A3:A4" },
        WorkbookAdvancedRowSpec.Create("Name", "Value", "Calc"),
        WorkbookAdvancedRowSpec.Create("{#lines}", string.Empty, string.Empty),
        WorkbookAdvancedRowSpec.Create(
            WorkbookCellSpec.Text("{name}"),
            WorkbookCellSpec.Text("{qty}"),
            WorkbookCellSpec.Formula("B3*2")),
        WorkbookAdvancedRowSpec.Create(
            WorkbookCellSpec.Text(string.Empty),
            WorkbookCellSpec.Text("{price}"),
            WorkbookCellSpec.Formula("B4*3")),
        WorkbookAdvancedRowSpec.Create("{/lines}", string.Empty, string.Empty),
        WorkbookAdvancedRowSpec.Create(
            WorkbookCellSpec.Text("Total"),
            WorkbookCellSpec.Text(string.Empty),
            WorkbookCellSpec.Formula("SUM(C3:C4)")));

    File.WriteAllBytes(templatePath, templateBytes);

    var engine = new XlsxTemplateEngine();
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

static void GeneratePresentationExample(
    string examplesRoot,
    string name,
    string json,
    string sampleCode,
    params PresentationSlideSpec[] slides)
{
    var dir = Path.Combine(examplesRoot, name);
    Directory.CreateDirectory(dir);

    var templatePath = Path.Combine(dir, "template.pptx");
    var dataPath = Path.Combine(dir, "data.json");
    var outputPath = Path.Combine(dir, "output.pptx");
    var codePath = Path.Combine(dir, "example.cs");

    File.WriteAllText(dataPath, json.Trim() + Environment.NewLine);
    File.WriteAllText(codePath, sampleCode.Trim() + Environment.NewLine);

    var templateBytes = CreatePresentationTemplate(slides);
    File.WriteAllBytes(templatePath, templateBytes);

    var engine = new PptxTemplateEngine();
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

static TableCell CellWithSplitRuns(params string[] pieces)
{
    var paragraph = new Paragraph();
    foreach (var piece in pieces)
    {
        paragraph.Append(new Run(new Text(piece)));
    }

    return new TableCell(paragraph);
}

static byte[] CreateWorkbookTemplate(params WorkbookRowSpec[] rows)
{
    using var stream = new MemoryStream();
    using (var document = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook, true))
    {
        var workbookPart = document.AddWorkbookPart();
        workbookPart.Workbook = new S.Workbook();

        var sharedStringPart = workbookPart.AddNewPart<SharedStringTablePart>();
        sharedStringPart.SharedStringTable = new S.SharedStringTable();

        var stylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
        stylesPart.Stylesheet = CreateWorkbookStylesheet();
        stylesPart.Stylesheet.Save();

        var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
        var sheetData = new S.SheetData();

        for (var rowIndex = 0; rowIndex < rows.Length; rowIndex++)
        {
            var row = new S.Row { RowIndex = (uint)(rowIndex + 1) };
            for (var columnIndex = 0; columnIndex < rows[rowIndex].Values.Length; columnIndex++)
            {
                var cell = CreateWorkbookSharedStringCell(
                    workbookPart,
                    rows[rowIndex].Values[columnIndex],
                    GetWorkbookCellReference(columnIndex + 1, rowIndex + 1));

                var styleIndex = rows[rowIndex].GetStyleIndex(columnIndex);
                if (styleIndex.HasValue)
                {
                    cell.StyleIndex = styleIndex.Value;
                }

                row.Append(cell);
            }

            sheetData.Append(row);
        }

        var lastReference = rows.Length == 0
            ? "A1"
            : GetWorkbookCellReference(rows[0].Values.Length, rows.Length);

        worksheetPart.Worksheet = new S.Worksheet(
            new S.SheetDimension { Reference = "A1:" + lastReference },
            sheetData);
        worksheetPart.Worksheet.Save();

        workbookPart.Workbook.Append(
            new S.Sheets(
                new S.Sheet
                {
                    Id = workbookPart.GetIdOfPart(worksheetPart),
                    SheetId = 1U,
                    Name = "Report"
                }));
        workbookPart.Workbook.Save();
    }

    return stream.ToArray();
}

static byte[] CreateAdvancedWorkbookTemplate(IReadOnlyList<string> mergedRanges, params WorkbookAdvancedRowSpec[] rows)
{
    using var stream = new MemoryStream();
    using (var document = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook, true))
    {
        var workbookPart = document.AddWorkbookPart();
        workbookPart.Workbook = new S.Workbook();

        var sharedStringPart = workbookPart.AddNewPart<SharedStringTablePart>();
        sharedStringPart.SharedStringTable = new S.SharedStringTable();

        var stylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
        stylesPart.Stylesheet = CreateWorkbookStylesheet();
        stylesPart.Stylesheet.Save();

        var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
        var sheetData = new S.SheetData();

        for (var rowIndex = 0; rowIndex < rows.Length; rowIndex++)
        {
            var row = new S.Row { RowIndex = (uint)(rowIndex + 1) };
            for (var columnIndex = 0; columnIndex < rows[rowIndex].Cells.Length; columnIndex++)
            {
                var cellSpec = rows[rowIndex].Cells[columnIndex];
                var cellReference = GetWorkbookCellReference(columnIndex + 1, rowIndex + 1);
                var cell = cellSpec.IsFormula
                    ? CreateWorkbookFormulaCell(cellSpec.Value, cellReference)
                    : CreateWorkbookSharedStringCell(workbookPart, cellSpec.Value, cellReference);

                if (cellSpec.StyleIndex.HasValue)
                {
                    cell.StyleIndex = cellSpec.StyleIndex.Value;
                }

                row.Append(cell);
            }

            sheetData.Append(row);
        }

        var lastReference = rows.Length == 0
            ? "A1"
            : GetWorkbookCellReference(rows[0].Cells.Length, rows.Length);

        var worksheet = new S.Worksheet(
            new S.SheetDimension { Reference = "A1:" + lastReference },
            sheetData);

        if (mergedRanges.Count > 0)
        {
            var mergeCells = new S.MergeCells();
            foreach (var mergedRange in mergedRanges)
            {
                mergeCells.Append(new S.MergeCell { Reference = mergedRange });
            }

            mergeCells.Count = (uint)mergedRanges.Count;
            worksheet.Append(mergeCells);
        }

        worksheetPart.Worksheet = worksheet;
        worksheetPart.Worksheet.Save();

        workbookPart.Workbook.Append(
            new S.Sheets(
                new S.Sheet
                {
                    Id = workbookPart.GetIdOfPart(worksheetPart),
                    SheetId = 1U,
                    Name = "Report"
                }));
        workbookPart.Workbook.Save();
    }

    return stream.ToArray();
}

static S.Stylesheet CreateWorkbookStylesheet()
{
    return new S.Stylesheet(
        new S.Fonts(new S.Font(), new S.Font(new S.Bold())),
        new S.Fills(new S.Fill(new S.PatternFill { PatternType = S.PatternValues.None }), new S.Fill(new S.PatternFill { PatternType = S.PatternValues.Gray125 })),
        new S.Borders(new S.Border()),
        new S.CellStyleFormats(new S.CellFormat()),
        new S.CellFormats(
            new S.CellFormat(),
            new S.CellFormat { FontId = 1U, FillId = 0U, BorderId = 0U, ApplyFont = true }));
}

static S.Cell CreateWorkbookSharedStringCell(WorkbookPart workbookPart, string text, string cellReference)
{
    var sharedStringTable = workbookPart.SharedStringTablePart!.SharedStringTable!;
    sharedStringTable.AppendChild(new S.SharedStringItem(new S.Text(text ?? string.Empty)));
    var index = sharedStringTable.Elements<S.SharedStringItem>().Count() - 1;

    return new S.Cell
    {
        CellReference = cellReference,
        DataType = S.CellValues.SharedString,
        CellValue = new S.CellValue(index.ToString(System.Globalization.CultureInfo.InvariantCulture))
    };
}

static S.Cell CreateWorkbookFormulaCell(string formula, string cellReference)
{
    return new S.Cell
    {
        CellReference = cellReference,
        CellFormula = new S.CellFormula(formula),
        CellValue = new S.CellValue(string.Empty)
    };
}

static byte[] CreatePresentationTemplate(params PresentationSlideSpec[] slides)
{
    using var stream = new MemoryStream();
    using (var document = PresentationDocument.Create(stream, PresentationDocumentType.Presentation, true))
    {
        var presentationPart = document.AddPresentationPart();
        presentationPart.Presentation = new P.Presentation();

        var slideMasterPart = presentationPart.AddNewPart<SlideMasterPart>("rId1");
        var themePart = slideMasterPart.AddNewPart<ThemePart>("rId5");
        themePart.Theme = CreatePresentationTheme();
        themePart.Theme.Save();

        var slideLayoutPart = slideMasterPart.AddNewPart<SlideLayoutPart>("rId1");
        slideLayoutPart.SlideLayout = new P.SlideLayout(
            new P.CommonSlideData(CreatePresentationShapeTree()),
            new P.ColorMapOverride(new A.MasterColorMapping()));
        slideLayoutPart.SlideLayout.Save();

        slideMasterPart.SlideMaster = new P.SlideMaster(
            new P.CommonSlideData(CreatePresentationShapeTree()),
            new P.ColorMap
            {
                Background1 = A.ColorSchemeIndexValues.Light1,
                Text1 = A.ColorSchemeIndexValues.Dark1,
                Background2 = A.ColorSchemeIndexValues.Light2,
                Text2 = A.ColorSchemeIndexValues.Dark2,
                Accent1 = A.ColorSchemeIndexValues.Accent1,
                Accent2 = A.ColorSchemeIndexValues.Accent2,
                Accent3 = A.ColorSchemeIndexValues.Accent3,
                Accent4 = A.ColorSchemeIndexValues.Accent4,
                Accent5 = A.ColorSchemeIndexValues.Accent5,
                Accent6 = A.ColorSchemeIndexValues.Accent6,
                Hyperlink = A.ColorSchemeIndexValues.Hyperlink,
                FollowedHyperlink = A.ColorSchemeIndexValues.FollowedHyperlink
            },
            new P.SlideLayoutIdList(new P.SlideLayoutId { Id = 2147483649U, RelationshipId = "rId1" }),
            new P.TextStyles(new P.TitleStyle(), new P.BodyStyle(), new P.OtherStyle()));
        slideMasterPart.SlideMaster.Save();

        presentationPart.Presentation.SlideMasterIdList = new P.SlideMasterIdList(
            new P.SlideMasterId { Id = 2147483648U, RelationshipId = "rId1" });

        var slideIdList = new P.SlideIdList();
        for (var index = 0; index < slides.Length; index++)
        {
            var relationshipId = "rIdSlide" + (index + 1).ToString();
            var slidePart = presentationPart.AddNewPart<SlidePart>(relationshipId);
            slidePart.AddPart(slideLayoutPart, "rId1");
            slidePart.Slide = new P.Slide(
                new P.CommonSlideData(CreatePresentationShapeTree(slides[index].Texts)),
                new P.ColorMapOverride(new A.MasterColorMapping()));
            slidePart.Slide.Save();

            slideIdList.Append(new P.SlideId
            {
                Id = (uint)(256 + index),
                RelationshipId = relationshipId
            });
        }

        presentationPart.Presentation.SlideIdList = slideIdList;
        presentationPart.Presentation.SlideSize = new P.SlideSize { Cx = 9144000, Cy = 6858000 };
        presentationPart.Presentation.NotesSize = new P.NotesSize { Cx = 6858000, Cy = 9144000 };
        presentationPart.Presentation.Save();
    }

    return stream.ToArray();
}

static P.ShapeTree CreatePresentationShapeTree(params string[] texts)
{
    var shapeTree = new P.ShapeTree(
        new P.NonVisualGroupShapeProperties(
            new P.NonVisualDrawingProperties { Id = 1U, Name = string.Empty },
            new P.NonVisualGroupShapeDrawingProperties(),
            new P.ApplicationNonVisualDrawingProperties()),
        new P.GroupShapeProperties(new A.TransformGroup()));

    for (var index = 0; index < texts.Length; index++)
    {
        shapeTree.Append(CreatePresentationTextShape((uint)(index + 2), "TextBox " + (index + 1), texts[index], index));
    }

    return shapeTree;
}

static P.Shape CreatePresentationTextShape(uint id, string name, string text, int order)
{
    var offsetY = 400000L + (order * 900000L);
    return new P.Shape(
        new P.NonVisualShapeProperties(
            new P.NonVisualDrawingProperties { Id = id, Name = name },
            new P.NonVisualShapeDrawingProperties(new A.ShapeLocks { NoGrouping = true }),
            new P.ApplicationNonVisualDrawingProperties()),
        new P.ShapeProperties(
            new A.Transform2D(
                new A.Offset { X = 400000L, Y = offsetY },
                new A.Extents { Cx = 7600000L, Cy = 700000L }),
            new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }),
        new P.TextBody(
            new A.BodyProperties(),
            new A.ListStyle(),
            new A.Paragraph(
                new A.Run(new A.Text(text)),
                new A.EndParagraphRunProperties())));
}

static A.Theme CreatePresentationTheme()
{
    var theme = new A.Theme { Name = "NDocxTemplater Theme" };
    theme.Append(new A.ThemeElements(
        new A.ColorScheme(
            new A.Dark1Color(new A.SystemColor { Val = A.SystemColorValues.WindowText, LastColor = "000000" }),
            new A.Light1Color(new A.SystemColor { Val = A.SystemColorValues.Window, LastColor = "FFFFFF" }),
            new A.Dark2Color(new A.RgbColorModelHex { Val = "1F497D" }),
            new A.Light2Color(new A.RgbColorModelHex { Val = "EEECE1" }),
            new A.Accent1Color(new A.RgbColorModelHex { Val = "4F81BD" }),
            new A.Accent2Color(new A.RgbColorModelHex { Val = "C0504D" }),
            new A.Accent3Color(new A.RgbColorModelHex { Val = "9BBB59" }),
            new A.Accent4Color(new A.RgbColorModelHex { Val = "8064A2" }),
            new A.Accent5Color(new A.RgbColorModelHex { Val = "4BACC6" }),
            new A.Accent6Color(new A.RgbColorModelHex { Val = "F79646" }),
            new A.Hyperlink(new A.RgbColorModelHex { Val = "0000FF" }),
            new A.FollowedHyperlinkColor(new A.RgbColorModelHex { Val = "800080" }))
        { Name = "Office" },
        new A.FontScheme(
            new A.MajorFont(
                new A.LatinFont { Typeface = "Calibri" },
                new A.EastAsianFont { Typeface = string.Empty },
                new A.ComplexScriptFont { Typeface = string.Empty }),
            new A.MinorFont(
                new A.LatinFont { Typeface = "Calibri" },
                new A.EastAsianFont { Typeface = string.Empty },
                new A.ComplexScriptFont { Typeface = string.Empty }))
        { Name = "Office" },
        new A.FormatScheme(
            new A.FillStyleList(
                new A.SolidFill(new A.SchemeColor { Val = A.SchemeColorValues.PhColor }),
                new A.GradientFill(
                    new A.GradientStopList(
                        new A.GradientStop(new A.SchemeColor(new A.Tint { Val = 50000 }, new A.SaturationModulation { Val = 300000 }) { Val = A.SchemeColorValues.PhColor }) { Position = 0 },
                        new A.GradientStop(new A.SchemeColor(new A.Tint { Val = 37000 }, new A.SaturationModulation { Val = 300000 }) { Val = A.SchemeColorValues.PhColor }) { Position = 35000 },
                        new A.GradientStop(new A.SchemeColor(new A.Tint { Val = 15000 }, new A.SaturationModulation { Val = 350000 }) { Val = A.SchemeColorValues.PhColor }) { Position = 100000 }),
                    new A.LinearGradientFill { Angle = 16200000, Scaled = true }),
                new A.GradientFill(
                    new A.GradientStopList(
                        new A.GradientStop(new A.SchemeColor(new A.Shade { Val = 51000 }, new A.SaturationModulation { Val = 130000 }) { Val = A.SchemeColorValues.PhColor }) { Position = 0 },
                        new A.GradientStop(new A.SchemeColor(new A.Shade { Val = 93000 }, new A.SaturationModulation { Val = 130000 }) { Val = A.SchemeColorValues.PhColor }) { Position = 80000 },
                        new A.GradientStop(new A.SchemeColor(new A.Shade { Val = 94000 }, new A.SaturationModulation { Val = 135000 }) { Val = A.SchemeColorValues.PhColor }) { Position = 100000 }),
                    new A.LinearGradientFill { Angle = 16200000, Scaled = false })),
            new A.LineStyleList(
                new A.Outline(
                    new A.SolidFill(new A.SchemeColor { Val = A.SchemeColorValues.PhColor }),
                    new A.PresetDash { Val = A.PresetLineDashValues.Solid })
                { Width = 9525, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center },
                new A.Outline(
                    new A.SolidFill(new A.SchemeColor { Val = A.SchemeColorValues.PhColor }),
                    new A.PresetDash { Val = A.PresetLineDashValues.Solid })
                { Width = 25400, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center },
                new A.Outline(
                    new A.SolidFill(new A.SchemeColor { Val = A.SchemeColorValues.PhColor }),
                    new A.PresetDash { Val = A.PresetLineDashValues.Solid })
                { Width = 38100, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center }),
            new A.EffectStyleList(
                new A.EffectStyle(new A.EffectList()),
                new A.EffectStyle(new A.EffectList()),
                new A.EffectStyle(
                    new A.EffectList(
                        new A.OuterShadow
                        {
                            BlurRadius = 40000,
                            Distance = 20000,
                            Direction = 5400000,
                            RotateWithShape = false
                        }))),
            new A.BackgroundFillStyleList(
                new A.SolidFill(new A.SchemeColor { Val = A.SchemeColorValues.PhColor }),
                new A.GradientFill(
                    new A.GradientStopList(
                        new A.GradientStop(new A.SchemeColor(new A.Tint { Val = 40000 }, new A.SaturationModulation { Val = 350000 }) { Val = A.SchemeColorValues.PhColor }) { Position = 0 },
                        new A.GradientStop(new A.SchemeColor(new A.Tint { Val = 45000 }, new A.Shade { Val = 99000 }, new A.SaturationModulation { Val = 350000 }) { Val = A.SchemeColorValues.PhColor }) { Position = 40000 },
                        new A.GradientStop(new A.SchemeColor(new A.Shade { Val = 20000 }, new A.SaturationModulation { Val = 255000 }) { Val = A.SchemeColorValues.PhColor }) { Position = 100000 }),
                    new A.PathGradientFill { Path = A.PathShadeValues.Circle }),
                new A.GradientFill(
                    new A.GradientStopList(
                        new A.GradientStop(new A.SchemeColor(new A.Tint { Val = 80000 }, new A.SaturationModulation { Val = 300000 }) { Val = A.SchemeColorValues.PhColor }) { Position = 0 },
                        new A.GradientStop(new A.SchemeColor(new A.Shade { Val = 30000 }, new A.SaturationModulation { Val = 200000 }) { Val = A.SchemeColorValues.PhColor }) { Position = 100000 }),
                    new A.PathGradientFill { Path = A.PathShadeValues.Circle })))
        { Name = "Office" }));
    return theme;
}

static string GetWorkbookCellReference(int columnIndex, int rowIndex)
{
    return GetWorkbookColumnName(columnIndex) + rowIndex.ToString();
}

static string GetWorkbookColumnName(int columnIndex)
{
    var letters = new Stack<char>();
    var index = columnIndex;
    while (index > 0)
    {
        index--;
        letters.Push((char)('A' + (index % 26)));
        index /= 26;
    }

    return new string(letters.ToArray());
}

readonly struct WorkbookRowSpec
{
    public WorkbookRowSpec(string[] values, uint?[] styleIndexes)
    {
        Values = values;
        StyleIndexes = styleIndexes;
    }

    public string[] Values { get; }

    public uint?[] StyleIndexes { get; }

    public static WorkbookRowSpec Create(params string[] values)
    {
        return new WorkbookRowSpec(values, new uint?[values.Length]);
    }

    public static WorkbookRowSpec Create(string first, string second, string third, uint styleIndex)
    {
        var values = new[] { first, second, third };
        return new WorkbookRowSpec(values, Enumerable.Repeat<uint?>(styleIndex, values.Length).ToArray());
    }

    public uint? GetStyleIndex(int columnIndex)
    {
        return columnIndex >= 0 && columnIndex < StyleIndexes.Length ? StyleIndexes[columnIndex] : null;
    }
}

readonly struct WorkbookAdvancedRowSpec
{
    public WorkbookAdvancedRowSpec(WorkbookCellSpec[] cells)
    {
        Cells = cells;
    }

    public WorkbookCellSpec[] Cells { get; }

    public static WorkbookAdvancedRowSpec Create(params string[] values)
    {
        return new WorkbookAdvancedRowSpec(values.Select(static value => WorkbookCellSpec.Text(value)).ToArray());
    }

    public static WorkbookAdvancedRowSpec Create(params WorkbookCellSpec[] cells)
    {
        return new WorkbookAdvancedRowSpec(cells);
    }
}

readonly struct WorkbookCellSpec
{
    public WorkbookCellSpec(string value, bool isFormula, uint? styleIndex = null)
    {
        Value = value;
        IsFormula = isFormula;
        StyleIndex = styleIndex;
    }

    public string Value { get; }

    public bool IsFormula { get; }

    public uint? StyleIndex { get; }

    public static WorkbookCellSpec Text(string value, uint? styleIndex = null)
    {
        return new WorkbookCellSpec(value, false, styleIndex);
    }

    public static WorkbookCellSpec Formula(string formula, uint? styleIndex = null)
    {
        return new WorkbookCellSpec(formula, true, styleIndex);
    }
}

readonly struct PresentationSlideSpec
{
    public PresentationSlideSpec(string[] texts)
    {
        Texts = texts;
    }

    public string[] Texts { get; }

    public static PresentationSlideSpec Create(params string[] texts)
    {
        return new PresentationSlideSpec(texts);
    }
}
