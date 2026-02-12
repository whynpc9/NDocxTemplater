using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace NDocxTemplater.Tests;

public class DocxTemplateEngineTests
{
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
}
