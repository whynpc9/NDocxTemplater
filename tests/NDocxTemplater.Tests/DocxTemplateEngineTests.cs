using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Xunit;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;

namespace NDocxTemplater.Tests;

public class DocxTemplateEngineTests
{
    private const string TinyPngDataUri =
        "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO8B9pYAAAAASUVORK5CYII=";

    private readonly DocxTemplateEngine _engine = new DocxTemplateEngine();

    [Fact]
    public void Render_ReplacesBasicTags_WithJsonPathAndIndex()
    {
        var template = CreateTemplate(
            Paragraph("Patient: {patient.name}"),
            Paragraph("First code: {report.items[0].code}"),
            Paragraph("Last value: {report.items[1].value}"));

        const string json = @"{
  ""patient"": { ""name"": ""Alice"" },
  ""report"": {
    ""items"": [
      { ""code"": ""A1"", ""value"": 10 },
      { ""code"": ""B2"", ""value"": 25 }
    ]
  }
}";

        var output = _engine.Render(template, json);
        var lines = ReadBodyParagraphTexts(output);

        Assert.Contains("Patient: Alice", lines);
        Assert.Contains("First code: A1", lines);
        Assert.Contains("Last value: 25", lines);
    }

    [Fact]
    public void Render_EvaluatesConditionalBlocks()
    {
        var template = CreateTemplate(
            Paragraph("Header"),
            Paragraph("{?flags.showVip}"),
            Paragraph("VIP Section"),
            Paragraph("{/?flags.showVip}"),
            Paragraph("Footer"));

        const string trueJson = @"{ ""flags"": { ""showVip"": true } }";
        const string falseJson = @"{ ""flags"": { ""showVip"": false } }";

        var whenTrue = ReadBodyParagraphTexts(_engine.Render(template, trueJson));
        var whenFalse = ReadBodyParagraphTexts(_engine.Render(template, falseJson));

        Assert.Contains("VIP Section", whenTrue);
        Assert.DoesNotContain("VIP Section", whenFalse);
        Assert.DoesNotContain("{?flags.showVip}", whenTrue);
        Assert.DoesNotContain("{/?flags.showVip}", whenTrue);
    }

    [Fact]
    public void Render_EvaluatesLoopBlocks()
    {
        var template = CreateTemplate(
            Paragraph("Items"),
            Paragraph("{#orders}"),
            Paragraph("- {id}: {amount}"),
            Paragraph("{/orders}"));

        const string json = @"{
  ""orders"": [
    { ""id"": ""ORD-1"", ""amount"": 120.5 },
    { ""id"": ""ORD-2"", ""amount"": 80 }
  ]
}";

        var lines = ReadBodyParagraphTexts(_engine.Render(template, json));

        Assert.Contains("- ORD-1: 120.5", lines);
        Assert.Contains("- ORD-2: 80", lines);
    }

    [Fact]
    public void Render_MapsJsonListToTableRows()
    {
        var template = CreateTemplate(CreateInvoiceTableTemplate());

        const string json = @"{
  ""invoice"": {
    ""lines"": [
      { ""name"": ""Apple"", ""qty"": 2 },
      { ""name"": ""Banana"", ""qty"": 3 }
    ]
  }
}";

        var output = _engine.Render(template, json);
        var rows = ReadFirstTableRows(output);

        Assert.Equal(3, rows.Count);
        Assert.Equal(new[] { "Item", "Qty" }, rows[0]);
        Assert.Equal(new[] { "Apple", "2" }, rows[1]);
        Assert.Equal(new[] { "Banana", "3" }, rows[2]);
    }

    [Fact]
    public void Render_SupportsSortTakeCountAndFormattingExtensions()
    {
        var template = CreateTemplate(
            Paragraph("Orders count: {orders|count}"),
            Paragraph("Report date: {reportDate|format:date:yyyy-MM-dd}"),
            Paragraph("{#orders|sort:amount:desc|take:2}"),
            Paragraph("{id} -> {amount|format:number:0.00}"),
            Paragraph("{/orders|sort:amount:desc|take:2}"));

        const string json = @"{
  ""reportDate"": ""2026-02-10T16:45:30Z"",
  ""orders"": [
    { ""id"": ""ORD-001"", ""amount"": 12.5 },
    { ""id"": ""ORD-002"", ""amount"": 100 },
    { ""id"": ""ORD-003"", ""amount"": 66.2 }
  ]
}";

        var lines = ReadBodyParagraphTexts(_engine.Render(template, json));

        Assert.Contains("Orders count: 3", lines);
        Assert.Contains("Report date: 2026-02-10", lines);

        var renderedOrderLines = lines
            .Where(static line => line.StartsWith("ORD-", StringComparison.Ordinal))
            .ToArray();

        Assert.Equal(new[]
        {
            "ORD-002 -> 100.00",
            "ORD-003 -> 66.20"
        }, renderedOrderLines);
    }

    [Fact]
    public void Render_SupportsInlineFriendlyAggregateExpressions_ForNarrativeParagraphs()
    {
        var template = CreateTemplate(
            Paragraph(
                "统计数据包括了从{financeMonthly|sort:month:asc|first|get:month|format:date:yyyy年M月}到{financeMonthly|sort:month:asc|last|get:month|format:date:yyyy年M月}的财务数据，其中营收最高的是{financeMonthly|maxby:revenue|get:month|format:date:M月}，营收为{financeMonthly|maxby:revenue|get:revenue|format:number:#,##0}元"),
            Paragraph(
                "在这些机构的对比数据中，其中营收最高的为{institutions|maxby:revenue|get:name}，收入为{institutions|maxby:revenue|get:revenue|format:number:#,##0}元，营收最低的为{institutions|minby:revenue|get:name}，收入为{institutions|minby:revenue|get:revenue|format:number:#,##0}元"));

        const string json = @"{
  ""financeMonthly"": [
    { ""month"": ""2025-03-01"", ""revenue"": 90000 },
    { ""month"": ""2025-01-01"", ""revenue"": 70000 },
    { ""month"": ""2025-07-01"", ""revenue"": 85000 },
    { ""month"": ""2025-05-01"", ""revenue"": 100000 }
  ],
  ""institutions"": [
    { ""name"": ""机构C"", ""revenue"": 650000 },
    { ""name"": ""机构A"", ""revenue"": 1000000 },
    { ""name"": ""机构Z"", ""revenue"": 100000 }
  ]
}";

        var lines = ReadBodyParagraphTexts(_engine.Render(template, json));

        Assert.Contains(
            "统计数据包括了从2025年1月到2025年7月的财务数据，其中营收最高的是5月，营收为100,000元",
            lines);
        Assert.Contains(
            "在这些机构的对比数据中，其中营收最高的为机构A，收入为1,000,000元，营收最低的为机构Z，收入为100,000元",
            lines);
    }

    [Fact]
    public void Render_SupportsNthAndAtExpressions_ForRankedInlineNarrative()
    {
        var template = CreateTemplate(
            Paragraph(
                "前10名机构中，第3名为{institutions|sort:revenue:desc|take:10|nth:3|get:name}，收入为{institutions|sort:revenue:desc|take:10|nth:3|get:revenue|format:number:#,##0}元；前10名末位为{institutions|sort:revenue:desc|take:10|at:-1|get:name}。"));

        const string json = @"{
  ""institutions"": [
    { ""name"": ""机构A"", ""revenue"": 1000000 },
    { ""name"": ""机构B"", ""revenue"": 920000 },
    { ""name"": ""机构C"", ""revenue"": 880000 },
    { ""name"": ""机构D"", ""revenue"": 860000 },
    { ""name"": ""机构E"", ""revenue"": 840000 },
    { ""name"": ""机构F"", ""revenue"": 820000 },
    { ""name"": ""机构G"", ""revenue"": 800000 },
    { ""name"": ""机构H"", ""revenue"": 780000 },
    { ""name"": ""机构I"", ""revenue"": 760000 },
    { ""name"": ""机构J"", ""revenue"": 740000 },
    { ""name"": ""机构K"", ""revenue"": 100000 }
  ]
}";

        var lines = ReadBodyParagraphTexts(_engine.Render(template, json));

        Assert.Contains(
            "前10名机构中，第3名为机构C，收入为880,000元；前10名末位为机构J。",
            lines);
    }

    [Fact]
    public void Render_SupportsInlineCountConditionalAndPercentPermilleFormatting()
    {
        var template = CreateTemplate(
            Paragraph(
                "本次样本共{institutions|count}家机构，状态：{flags.includeRates|if:包含比率指标:不包含比率指标}，环比增长率{metrics.growthRate|format:percent:0.00}，坏账率{metrics.badDebtRate|format:permille:0.00}。"),
            Paragraph(
                "备用写法（number pattern）：{metrics.growthRate|format:number:0.00%} / {metrics.badDebtRate|format:number:0.00‰}"));

        const string json = @"{
  ""flags"": { ""includeRates"": true },
  ""metrics"": {
    ""growthRate"": 0.0123,
    ""badDebtRate"": 0.0045
  },
  ""institutions"": [
    { ""name"": ""机构A"" },
    { ""name"": ""机构B"" },
    { ""name"": ""机构C"" }
  ]
}";

        var lines = ReadBodyParagraphTexts(_engine.Render(template, json));

        Assert.Contains(
            "本次样本共3家机构，状态：包含比率指标，环比增长率1.23%，坏账率4.50‰。",
            lines);
        Assert.Contains(
            "备用写法（number pattern）：1.23% / 4.50‰",
            lines);
    }

    [Fact]
    public void Render_RendersInlineImageTag_FromDataUri()
    {
        var template = CreateTemplate(
            Paragraph("Report logo"),
            Paragraph("{%logo}"));

        var json = @"{
  ""logo"": {
    ""src"": """ + TinyPngDataUri + @""",
    ""width"": 32,
    ""height"": 16
  }
}";

        var output = _engine.Render(template, json);

        using (var stream = new MemoryStream(output))
        using (var document = WordprocessingDocument.Open(stream, false))
        {
            var drawing = document.MainDocumentPart!.Document.Body!.Descendants<Drawing>().Single();
            var extent = drawing.Descendants<DW.Extent>().Single();

            Assert.Equal(32 * 9525L, extent.Cx!.Value);
            Assert.Equal(16 * 9525L, extent.Cy!.Value);
        }
    }

    [Fact]
    public void Render_RendersBlockImageTag_Centered()
    {
        var template = CreateTemplate(Paragraph("{%%cover}"));

        var json = @"{
  ""cover"": {
    ""src"": """ + TinyPngDataUri + @""",
    ""width"": 24,
    ""height"": 24
  }
}";

        var output = _engine.Render(template, json);

        using (var stream = new MemoryStream(output))
        using (var document = WordprocessingDocument.Open(stream, false))
        {
            var paragraph = document.MainDocumentPart!.Document.Body!.Elements<Paragraph>().Single();
            var drawing = paragraph.Descendants<Drawing>().Single();
            var extent = drawing.Descendants<DW.Extent>().Single();
            var alignment = paragraph.ParagraphProperties?.Justification?.Val?.Value;

            Assert.Equal(24 * 9525L, extent.Cx!.Value);
            Assert.Equal(24 * 9525L, extent.Cy!.Value);
            Assert.Equal(JustificationValues.Center, alignment);
        }
    }

    [Fact]
    public void Render_RendersImagesInsideLoopBlocks()
    {
        var template = CreateTemplate(
            Paragraph("{#gallery}"),
            Paragraph("{%%photo}"),
            Paragraph("{/gallery}"));

        var json = @"{
  ""gallery"": [
    { ""photo"": { ""src"": """ + TinyPngDataUri + @""", ""width"": 18, ""height"": 18 } },
    { ""photo"": { ""src"": """ + TinyPngDataUri + @""", ""width"": 20, ""height"": 20 } }
  ]
}";

        var output = _engine.Render(template, json);

        using (var stream = new MemoryStream(output))
        using (var document = WordprocessingDocument.Open(stream, false))
        {
            var drawings = document.MainDocumentPart!.Document.Body!.Descendants<Drawing>().ToList();
            Assert.Equal(2, drawings.Count);
        }
    }

    [Fact]
    public void Render_RendersRealPngFromFilePathAndDataUri_WithAspectRatioScaling()
    {
        var imagePath = GetTestAssetPath("real-chart.png");
        var imageBytes = File.ReadAllBytes(imagePath);
        var imageDataUri = "data:image/png;base64," + Convert.ToBase64String(imageBytes);

        var template = CreateTemplate(
            Paragraph("File path image"),
            Paragraph("{%fileImage}"),
            Paragraph("Data URI image"),
            Paragraph("{%dataUriImage}"),
            Paragraph("Fit box image"),
            Paragraph("{%fitBoxImage}"));

        var json = @"{
  ""fileImage"": {
    ""src"": """ + EscapeJsonString(imagePath) + @""",
    ""maxWidth"": 376,
    ""preserveAspectRatio"": true
  },
  ""dataUriImage"": {
    ""src"": """ + imageDataUri + @""",
    ""scale"": 0.25,
    ""preserveAspectRatio"": true
  },
  ""fitBoxImage"": {
    ""src"": """ + EscapeJsonString(imagePath) + @""",
    ""width"": 420,
    ""height"": 260,
    ""preserveAspectRatio"": true
  }
}";

        var output = _engine.Render(template, json);

        using (var stream = new MemoryStream(output))
        using (var document = WordprocessingDocument.Open(stream, false))
        {
            var drawings = document.MainDocumentPart!.Document.Body!.Descendants<Drawing>().ToList();
            Assert.Equal(3, drawings.Count);

            var extents = drawings
                .Select(static drawing => drawing.Descendants<DW.Extent>().Single())
                .Select(static extent => (extent.Cx!.Value, extent.Cy!.Value))
                .ToArray();

            Assert.Equal((376 * 9525L, 339 * 9525L), extents[0]);
            Assert.Equal((376 * 9525L, 339 * 9525L), extents[1]);
            Assert.Equal((288 * 9525L, 260 * 9525L), extents[2]);

            var embeddedImages = document.MainDocumentPart.ImageParts
                .Select(static part =>
                {
                    using var imageStream = part.GetStream();
                    using var copy = new MemoryStream();
                    imageStream.CopyTo(copy);
                    return copy.ToArray();
                })
                .ToList();

            Assert.Equal(3, embeddedImages.Count);
            Assert.All(embeddedImages, bytes => Assert.Equal(imageBytes, bytes));
        }
    }

    [Fact]
    public void Render_FormatsDateExpressionInTable_WhenTagIsSplitAcrossRuns()
    {
        var template = CreateTemplate(
            new Table(
                TableRow(Cell("Name"), Cell("Created")),
                TableRow(Cell("{#rows}"), Cell(string.Empty)),
                TableRow(
                    Cell("{name}"),
                    CellWithSplitRuns("{createdAt|for", "mat:date:yyyy-MM-", "dd}")),
                TableRow(Cell("{/rows}"), Cell(string.Empty))));

        const string json = @"{
  ""rows"": [
    { ""name"": ""A"", ""createdAt"": ""2026-02-24T10:11:12Z"" }
  ]
}";

        var output = _engine.Render(template, json);
        var rows = ReadFirstTableRows(output);

        Assert.Equal(new[] { "A", "2026-02-24" }, rows[1]);
    }

    private static byte[] CreateTemplate(params OpenXmlElement[] bodyElements)
    {
        using (var stream = new MemoryStream())
        {
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
    }

    private static Table CreateInvoiceTableTemplate()
    {
        return new Table(
            TableRow(Cell("Item"), Cell("Qty")),
            TableRow(Cell("{#invoice.lines}"), Cell(string.Empty)),
            TableRow(Cell("{name}"), Cell("{qty}")),
            TableRow(Cell("{/invoice.lines}"), Cell(string.Empty)));
    }

    private static Paragraph Paragraph(string text)
    {
        return new Paragraph(new Run(new Text(text)));
    }

    private static TableRow TableRow(params TableCell[] cells)
    {
        return new TableRow(cells);
    }

    private static TableCell Cell(string text)
    {
        return new TableCell(new Paragraph(new Run(new Text(text))));
    }

    private static TableCell CellWithSplitRuns(params string[] pieces)
    {
        var paragraph = new Paragraph();
        foreach (var piece in pieces)
        {
            paragraph.Append(new Run(new Text(piece)));
        }

        return new TableCell(paragraph);
    }

    private static IReadOnlyList<string> ReadBodyParagraphTexts(byte[] docx)
    {
        using (var stream = new MemoryStream(docx))
        using (var document = WordprocessingDocument.Open(stream, false))
        {
            return document.MainDocumentPart!.Document.Body!
                .Elements<Paragraph>()
                .Select(static paragraph => string.Concat(paragraph.Descendants<Text>().Select(static text => text.Text)))
                .Where(static line => !string.IsNullOrWhiteSpace(line))
                .ToArray();
        }
    }

    private static IReadOnlyList<string[]> ReadFirstTableRows(byte[] docx)
    {
        using (var stream = new MemoryStream(docx))
        using (var document = WordprocessingDocument.Open(stream, false))
        {
            var table = document.MainDocumentPart!.Document.Body!.Elements<Table>().First();

            return table.Elements<TableRow>()
                .Select(static row => row.Elements<TableCell>()
                    .Select(static cell => string.Concat(cell.Descendants<Text>().Select(static text => text.Text)))
                    .ToArray())
                .ToArray();
        }
    }

    private static string GetTestAssetPath(string fileName)
    {
        var repoRoot = FindRepoRoot(AppContext.BaseDirectory);
        return Path.Combine(repoRoot, "tests", "NDocxTemplater.Tests", "Assets", fileName);
    }

    private static string FindRepoRoot(string startPath)
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

    private static string EscapeJsonString(string text)
    {
        return text
            .Replace("\\", "\\\\")
            .Replace("\"", "\\\"");
    }
}
