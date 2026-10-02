using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Drawing.Spreadsheet;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using SixLabors.ImageSharp;
using SixLabors.ImageSharp.PixelFormats;
using Xunit;

namespace NDocxTemplater.Tests;

public class XlsxTemplateEngineTests
{
    private readonly XlsxTemplateEngine _engine = new XlsxTemplateEngine();

    [Theory]
    [InlineData("A0", false, 0U)]
    [InlineData("A1", true, 1U)]
    [InlineData("XFD1048576", true, SpreadsheetRowBounds.Maximum)]
    [InlineData("A1048577", false, 0U)]
    [InlineData("A4294967295", false, 0U)]
    public void SpreadsheetCellReference_TryParse_EnforcesXlsxRowBounds(string reference, bool expectedResult, uint expectedRow)
    {
        var result = SpreadsheetCellReference.TryParse(reference, out var parsed);

        Assert.Equal(expectedResult, result);
        Assert.Equal(expectedRow, parsed.RowIndex);
    }

    [Fact]
    public void SpreadsheetRowMapping_ResolvesMaximumRowWithoutWraparound()
    {
        const string scopeId = "scope";
        var renderedRow = new RenderedSpreadsheetRow(
            new Row { RowIndex = SpreadsheetRowBounds.Maximum },
            SpreadsheetRowBounds.Maximum,
            scopeId)
        {
            TargetRowIndex = 42U
        };
        var mapping = new SpreadsheetRowMapping(new[] { renderedRow }, "root");

        Assert.True(mapping.TryResolveRangeInScope(
            scopeId,
            SpreadsheetRowBounds.Maximum,
            SpreadsheetRowBounds.Maximum,
            out var scopedStart,
            out var scopedEnd));
        Assert.Equal(42U, scopedStart);
        Assert.Equal(42U, scopedEnd);

        Assert.True(mapping.TryResolveRangeGlobal(
            SpreadsheetRowBounds.Maximum,
            SpreadsheetRowBounds.Maximum,
            out var globalStart,
            out var globalEnd));
        Assert.Equal(42U, globalStart);
        Assert.Equal(42U, globalEnd);
    }

    [Theory]
    [InlineData(0U)]
    [InlineData(1_048_577U)]
    [InlineData(uint.MaxValue)]
    public void SpreadsheetRowMapping_RejectsOutOfBoundsDirectCallers(uint invalidRow)
    {
        var renderedRow = new RenderedSpreadsheetRow(new Row { RowIndex = 1U }, 1U, "scope")
        {
            TargetRowIndex = 1U
        };
        var mapping = new SpreadsheetRowMapping(new[] { renderedRow }, "root");

        Assert.False(mapping.TryResolveRangeInScope("scope", 1U, invalidRow, out var scopedStart, out var scopedEnd));
        Assert.Equal(0U, scopedStart);
        Assert.Equal(0U, scopedEnd);

        Assert.False(mapping.TryResolveRangeGlobal(1U, invalidRow, out var globalStart, out var globalEnd));
        Assert.Equal(0U, globalStart);
        Assert.Equal(0U, globalEnd);
    }

    [Theory]
    [InlineData("A1:A1048577")]
    [InlineData("A1:A4294967295")]
    public void Render_PreservesMalformedOutOfBoundsRangeWithoutIteratingIt(string invalidRange)
    {
        var template = CreateWorkbookWithAutoFilterReference(invalidRange);

        var output = _engine.Render(template, "{}");

        using var stream = new MemoryStream(output);
        using var document = SpreadsheetDocument.Open(stream, false);
        var autoFilter = document.WorkbookPart!.WorksheetParts.First().Worksheet.GetFirstChild<AutoFilter>();
        Assert.Equal(invalidRange, autoFilter!.Reference!.Value);
    }

    [Fact]
    public void Render_ReplacesWorksheetCellTags_AndKeepsTypedNumericValues()
    {
        var template = CreateWorkbook(
            RowSpec.Create("Report", "{report.title}"),
            RowSpec.Create("Report date", "{reportDate|format:date:yyyy-MM-dd}"),
            RowSpec.Create("Orders count", "{orders|count}"));

        const string json = @"{
  ""report"": { ""title"": ""Sales Summary"" },
  ""reportDate"": ""2026-03-18T09:10:11Z"",
  ""orders"": [
    { ""id"": ""ORD-1"" },
    { ""id"": ""ORD-2"" },
    { ""id"": ""ORD-3"" }
  ]
}";

        var output = _engine.Render(template, json);
        var rows = ReadSheetRows(output);

        Assert.Equal(new[] { "Report", "Sales Summary" }, rows[0].Values);
        Assert.Equal(new[] { "Report date", "2026-03-18" }, rows[1].Values);
        Assert.Equal(new[] { "Orders count", "3" }, rows[2].Values);
    }

    [Fact]
    public void Render_CanKeepMissingWorksheetTags()
    {
        var template = CreateWorkbook(
            RowSpec.Create("Full tag", "{report.missing}"),
            RowSpec.Create("Inline tag", "Value: {report.missing}"));
        var warnings = new List<RenderWarning>();
        var options = new RenderOptions
        {
            MissingValueBehavior = MissingValueBehavior.KeepTag,
            WarningHandler = warnings.Add
        };

        const string json = @"{ ""report"": { ""title"": ""Sales Summary"" } }";

        var output = _engine.Render(template, json, options);
        var rows = ReadSheetRows(output);

        Assert.Equal(new[] { "Full tag", "{report.missing}" }, rows[0].Values);
        Assert.Equal(new[] { "Inline tag", "Value: {report.missing}" }, rows[1].Values);
        Assert.Equal(2, warnings.Count);
        Assert.All(warnings, warning => Assert.Equal("report.missing", warning.Expression));
    }

    [Fact]
    public void Render_MapsJsonListToWorksheetRows_AndPreservesTemplateCellStyles()
    {
        var template = CreateWorkbook(
            RowSpec.Create("ID", "Qty", "Amount"),
            RowSpec.Create("{#orders|sort:amount:desc|take:2}", string.Empty, string.Empty),
            RowSpec.Create(new[] { "{id}", "{qty}", "{amount|format:number:0.00}" }, styleIndex: 1U),
            RowSpec.Create("{/orders|sort:amount:desc|take:2}", string.Empty, string.Empty),
            RowSpec.Create("{?showFooter}", string.Empty, string.Empty),
            RowSpec.Create(new[] { "Count", "{orders|count}", string.Empty }, styleIndex: 1U),
            RowSpec.Create("{/?showFooter}", string.Empty, string.Empty));

        const string json = @"{
  ""showFooter"": true,
  ""orders"": [
    { ""id"": ""ORD-001"", ""qty"": 1, ""amount"": 12.5 },
    { ""id"": ""ORD-002"", ""qty"": 3, ""amount"": 100 },
    { ""id"": ""ORD-003"", ""qty"": 2, ""amount"": 66.2 }
  ]
}";

        var output = _engine.Render(template, json);
        var rows = ReadSheetRows(output);

        Assert.Equal(4, rows.Count);
        Assert.Equal(new[] { "ID", "Qty", "Amount" }, rows[0].Values);
        Assert.Equal(new[] { "ORD-002", "3", "100.00" }, rows[1].Values);
        Assert.Equal(new[] { "ORD-003", "2", "66.20" }, rows[2].Values);
        Assert.Equal(new[] { "Count", "3", string.Empty }, rows[3].Values);
        Assert.Equal(new uint?[] { 1U, 1U, 1U }, rows[1].StyleIndexes);
        Assert.Equal(new uint?[] { 1U, 1U, 1U }, rows[3].StyleIndexes);
    }

    [Fact]
    public void Render_InsertsWorksheetImagesAndBarcodes_FromCellPlaceholders()
    {
        var imagePath = GetTestAssetPath("real-chart.png");
        var imageBytes = File.ReadAllBytes(imagePath);
        var imageDataUri = "data:image/png;base64," + Convert.ToBase64String(imageBytes);
        var imageSize = GetImageSize(imageBytes);

        var template = CreateWorkbook(
            RowSpec.Create("Media", "Value"),
            RowSpec.Create("From path", "{%pathImage}"),
            RowSpec.Create("From data URI", "{%dataUriImage}"),
            RowSpec.Create("Barcode", "{%barcode:barcodes.ean13;type=ean13;width=220;height=80}"));

        var json = @"{
  ""pathImage"": {
    ""src"": """ + EscapeJsonString(imagePath) + @""",
    ""maxWidth"": 320,
    ""preserveAspectRatio"": true
  },
  ""dataUriImage"": {
    ""src"": """ + imageDataUri + @""",
    ""scale"": 0.2,
    ""preserveAspectRatio"": true
  },
  ""barcodes"": {
    ""ean13"": ""5901234123457""
  }
}";

        var output = _engine.Render(template, json);
        var rows = ReadSheetRows(output);

        Assert.Equal(string.Empty, rows[1].Values[1]);
        Assert.Equal(string.Empty, rows[2].Values[1]);
        Assert.Equal(string.Empty, rows[3].Values[1]);

        using (var stream = new MemoryStream(output))
        using (var document = SpreadsheetDocument.Open(stream, false))
        {
            var worksheetPart = document.WorkbookPart!.WorksheetParts.First();
            Assert.NotNull(worksheetPart.DrawingsPart);
            Assert.NotNull(worksheetPart.Worksheet.Elements<Drawing>().FirstOrDefault());

            var drawing = worksheetPart.DrawingsPart!.WorksheetDrawing!;
            var anchors = drawing.Elements<OneCellAnchor>().ToArray();
            Assert.Equal(3, anchors.Length);

            Assert.Equal("B2", AnchorToCell(anchors[0]));
            Assert.Equal("B3", AnchorToCell(anchors[1]));
            Assert.Equal("B4", AnchorToCell(anchors[2]));

            Assert.Equal(PixelsToEmu(320), anchors[0].Extent!.Cx!.Value);
            Assert.Equal(PixelsToEmu((int)Math.Round(imageSize.Height * (320d / imageSize.Width))), anchors[0].Extent!.Cy!.Value);

            Assert.Equal(PixelsToEmu((int)Math.Round(imageSize.Width * 0.2d)), anchors[1].Extent!.Cx!.Value);
            Assert.Equal(PixelsToEmu((int)Math.Round(imageSize.Height * 0.2d)), anchors[1].Extent!.Cy!.Value);

            var embeddedImages = worksheetPart.DrawingsPart.ImageParts
                .Select(part =>
                {
                    using var imageStream = part.GetStream();
                    using var copy = new MemoryStream();
                    imageStream.CopyTo(copy);
                    return copy.ToArray();
                })
                .ToArray();

            Assert.Equal(3, embeddedImages.Length);
            Assert.Equal(2, embeddedImages.Count(bytes => bytes.SequenceEqual(imageBytes)));
            Assert.Single(embeddedImages, bytes => !bytes.SequenceEqual(imageBytes));
            Assert.True(ContainsDarkPixels(embeddedImages.Single(bytes => !bytes.SequenceEqual(imageBytes))));
        }
    }

    [Theory]
    [InlineData("A1", "C1", "C1")]
    [InlineData(null, "C1", "C1")]
    [InlineData("C1", null, "D1")]
    [InlineData("C1", "not-a-reference", "D1")]
    public void Render_AnchorsMediaUsingSparseCellReferences(
        string? leadingCellReference,
        string? mediaCellReference,
        string expectedAnchor)
    {
        var imagePath = GetTestAssetPath("real-chart.png");
        var template = CreateWorkbookWithReferencedMediaCells(leadingCellReference, mediaCellReference);
        var json = "{\"image\":{\"src\":\"" + EscapeJsonString(imagePath) + "\"}}";

        var output = _engine.Render(template, json);

        using var stream = new MemoryStream(output);
        using var document = SpreadsheetDocument.Open(stream, false);
        var worksheetPart = document.WorkbookPart!.WorksheetParts.First();
        var anchor = Assert.Single(worksheetPart.DrawingsPart!.WorksheetDrawing!.Elements<OneCellAnchor>());
        Assert.Equal(expectedAnchor, AnchorToCell(anchor));
    }

    [Fact]
    public void Render_RepeatsMergedRanges_AndAdjustsFormulaReferences()
    {
        var template = CreateWorkbook(
            new[] { "A3:A4" },
            RowSpec.Create("Name", "Value", "Calc"),
            RowSpec.Create("{#lines}", string.Empty, string.Empty),
            RowSpec.Create(CellSpec.Text("{name}"), CellSpec.Text("{qty}"), CellSpec.Formula("B3*2")),
            RowSpec.Create(CellSpec.Text(string.Empty), CellSpec.Text("{price}"), CellSpec.Formula("B4*3")),
            RowSpec.Create("{/lines}", string.Empty, string.Empty),
            RowSpec.Create(CellSpec.Text("Total"), CellSpec.Text(string.Empty), CellSpec.Formula("SUM(C3:C4)")));

        const string json = @"{
  ""lines"": [
    { ""name"": ""Alpha"", ""qty"": 2, ""price"": 5 },
    { ""name"": ""Beta"", ""qty"": 3, ""price"": 7 }
  ]
}";

        var output = _engine.Render(template, json);
        var rows = ReadSheetRows(output);

        Assert.Equal(6, rows.Count);
        Assert.Equal(new[] { "Name", "Value", "Calc" }, rows[0].Values);
        Assert.Equal(new[] { "Alpha", "2", "B2*2" }, rows[1].Values);
        Assert.Equal(new[] { string.Empty, "5", "B3*3" }, rows[2].Values);
        Assert.Equal(new[] { "Beta", "3", "B4*2" }, rows[3].Values);
        Assert.Equal(new[] { string.Empty, "7", "B5*3" }, rows[4].Values);
        Assert.Equal(new[] { "Total", string.Empty, "SUM(C2:C5)" }, rows[5].Values);

        using (var stream = new MemoryStream(output))
        using (var document = SpreadsheetDocument.Open(stream, false))
        {
            var worksheet = document.WorkbookPart!.WorksheetParts.First().Worksheet;
            var mergeCells = worksheet.GetFirstChild<MergeCells>()!;
            var mergedRefs = mergeCells.Elements<MergeCell>()
                .Select(cell => cell.Reference!.Value!)
                .OrderBy(value => value, StringComparer.Ordinal)
                .ToArray();

            Assert.Equal(new[] { "A2:A3", "A4:A5" }, mergedRefs);
        }
    }

    [Fact]
    public void Render_UpdatesTablesFiltersDefinedNames_AndRemovesCalcChain()
    {
        var template = CreateWorkbookWithRangeMetadata();

        const string json = @"{
  ""lines"": [
    { ""name"": ""Alpha"", ""amount"": 10 },
    { ""name"": ""Beta"", ""amount"": 20 }
  ]
}";

        var output = _engine.Render(template, json);
        var rows = ReadSheetRows(output);

        Assert.Equal(4, rows.Count);
        Assert.Equal(new[] { "Name", "Amount", "Status" }, rows[0].Values);
        Assert.Equal(new[] { "Alpha", "10", "B2*2" }, rows[1].Values);
        Assert.Equal(new[] { "Beta", "20", "B3*2" }, rows[2].Values);
        Assert.Equal(new[] { "Total", string.Empty, "SUM(C2:C3)" }, rows[3].Values);

        using (var stream = new MemoryStream(output))
        using (var document = SpreadsheetDocument.Open(stream, false))
        {
            var workbookPart = document.WorkbookPart!;
            var worksheetPart = workbookPart.WorksheetParts.First();
            var table = worksheetPart.TableDefinitionParts.Single().Table;
            var definedName = workbookPart.Workbook.DefinedNames!.Elements<DefinedName>().Single();

            Assert.Equal("A1:C4", worksheetPart.Worksheet.GetFirstChild<AutoFilter>()!.Reference!.Value);
            Assert.Equal("A1:C4", table.Reference!.Value);
            Assert.Equal("A1:C4", table.AutoFilter!.Reference!.Value);
            Assert.Equal("'Report'!$A$1:$C$4", definedName.Text);
            Assert.Null(workbookPart.CalculationChainPart);
        }
    }

    [Theory]
    [InlineData(1, 4, "First")]
    [InlineData(0, 3, "First")]
    [InlineData(2, 4, "统计表")]
    [InlineData(1, 4, "2024")]
    public void Render_RewritesDefinedNamesOnlyWithTheirOwningSheetMapping(int firstCount, int secondCount, string firstSheetName)
    {
        using var templateStream = new MemoryStream();
        using (var document = SpreadsheetDocument.Create(templateStream, DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook, true))
        {
            var workbookPart = document.AddWorkbookPart();
            workbookPart.Workbook = new Workbook(new Sheets());
            workbookPart.AddNewPart<SharedStringTablePart>().SharedStringTable = new SharedStringTable();
            var sheetNames = new[] { firstSheetName, "O'Brien Data" };
            for (var index = 0; index < sheetNames.Length; index++)
            {
                var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                var expression = index == 0 ? "first" : "second";
                worksheetPart.Worksheet = new Worksheet(new SheetData(
                    CreateRow(workbookPart, 1, "Header"),
                    CreateRow(workbookPart, 2, "{#" + expression + "}"),
                    CreateRow(workbookPart, 3, "{name}"),
                    CreateRow(workbookPart, 4, "{/" + expression + "}"),
                    CreateRow(workbookPart, 5, "Footer")));
                workbookPart.Workbook.Sheets!.Append(new Sheet
                {
                    Id = workbookPart.GetIdOfPart(worksheetPart),
                    SheetId = (uint)(17 + index * 25),
                    Name = sheetNames[index]
                });
            }

            workbookPart.Workbook.Append(new DefinedNames(
                new DefinedName(firstSheetName + "!$A$1:$A$5") { Name = "FirstRange" },
                new DefinedName("'O''Brien Data'!$A$1:$A$5") { Name = "SecondRange" },
                new DefinedName("$A$1:$A$5") { Name = "LocalRange", LocalSheetId = 1U },
                new DefinedName(firstSheetName + "!$A$1:$A$5") { Name = "OtherSheet", LocalSheetId = 1U },
                new DefinedName(firstSheetName + "!$A$1:$A$5,'O''Brien Data'!$A$1:$A$5") { Name = "UnionRange" },
                new DefinedName("$A$1:$A$5") { Name = "UnqualifiedGlobal" },
                new DefinedName("'Unknown'!$A$1:$A$5") { Name = "UnknownSheet" },
                new DefinedName("[Other.xlsx]" + firstSheetName + "!$A$1:$A$5") { Name = "ExternalRange", LocalSheetId = 0U }));
            workbookPart.Workbook.Save();
        }

        var json = System.Text.Json.JsonSerializer.Serialize(new
        {
            first = Enumerable.Range(0, firstCount).Select(static index => new { name = "First " + index }).ToArray(),
            second = Enumerable.Range(0, secondCount).Select(static index => new { name = "Second " + index }).ToArray()
        });
        var output = _engine.Render(templateStream.ToArray(), json);
        using var outputStream = new MemoryStream(output);
        using var result = SpreadsheetDocument.Open(outputStream, false);
        var names = result.WorkbookPart!.Workbook.DefinedNames!.Elements<DefinedName>()
            .ToDictionary(static name => name.Name!.Value!, static name => name.Text);
        var firstRange = firstSheetName + "!$A$1:$A$" + (firstCount + 2);
        var secondRange = "'O''Brien Data'!$A$1:$A$" + (secondCount + 2);
        Assert.Equal(firstRange, names["FirstRange"]);
        Assert.Equal(secondRange, names["SecondRange"]);
        Assert.Equal("$A$1:$A$" + (secondCount + 2), names["LocalRange"]);
        Assert.Equal(firstRange, names["OtherSheet"]);
        Assert.Equal(firstRange + "," + secondRange, names["UnionRange"]);
        Assert.Equal("$A$1:$A$5", names["UnqualifiedGlobal"]);
        Assert.Equal("'Unknown'!$A$1:$A$5", names["UnknownSheet"]);
        Assert.Equal("[Other.xlsx]" + firstSheetName + "!$A$1:$A$5", names["ExternalRange"]);
    }

    private static byte[] CreateWorkbook(params RowSpec[] rows)
    {
        return CreateWorkbook(Array.Empty<string>(), rows);
    }

    private static byte[] CreateWorkbook(IReadOnlyList<string> mergedRanges, params RowSpec[] rows)
    {
        using (var stream = new MemoryStream())
        {
            using (var document = SpreadsheetDocument.Create(stream, DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook, true))
            {
                var workbookPart = document.AddWorkbookPart();
                workbookPart.Workbook = new Workbook();

                var sharedStringPart = workbookPart.AddNewPart<SharedStringTablePart>();
                sharedStringPart.SharedStringTable = new SharedStringTable();

                var stylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
                stylesPart.Stylesheet = CreateStylesheet();
                stylesPart.Stylesheet.Save();

                var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                var sheetData = new SheetData();

                for (var rowIndex = 0; rowIndex < rows.Length; rowIndex++)
                {
                    var row = new Row { RowIndex = (uint)(rowIndex + 1) };
                    for (var columnIndex = 0; columnIndex < rows[rowIndex].Cells.Length; columnIndex++)
                    {
                        var spec = rows[rowIndex].Cells[columnIndex];
                        var cellReference = GetCellReference(columnIndex + 1, rowIndex + 1);
                        var cell = spec.Kind == CellKind.Formula
                            ? CreateFormulaCell(spec.Value, cellReference)
                            : CreateSharedStringCell(workbookPart, spec.Value, cellReference);

                        if (spec.StyleIndex.HasValue)
                        {
                            cell.StyleIndex = spec.StyleIndex.Value;
                        }

                        row.Append(cell);
                    }

                    sheetData.Append(row);
                }

                var worksheet = new Worksheet(
                    new SheetDimension { Reference = rows.Length == 0 ? "A1" : "A1:" + GetCellReference(rows[0].Cells.Length, rows.Length) },
                    sheetData);

                if (mergedRanges.Count > 0)
                {
                    var mergeCells = new MergeCells();
                    foreach (var mergedRange in mergedRanges)
                    {
                        mergeCells.Append(new MergeCell { Reference = mergedRange });
                    }

                    mergeCells.Count = (uint)mergedRanges.Count;
                    worksheet.Append(mergeCells);
                }

                worksheetPart.Worksheet = worksheet;
                worksheetPart.Worksheet.Save();

                workbookPart.Workbook.Append(
                    new Sheets(
                        new Sheet
                        {
                            Id = workbookPart.GetIdOfPart(worksheetPart),
                            SheetId = 1U,
                            Name = "Report"
                        }));
                workbookPart.Workbook.Save();
            }

            return stream.ToArray();
        }
    }

    private static byte[] CreateWorkbookWithRangeMetadata()
    {
        using (var stream = new MemoryStream())
        {
            using (var document = SpreadsheetDocument.Create(stream, DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook, true))
            {
                var workbookPart = document.AddWorkbookPart();
                workbookPart.Workbook = new Workbook();

                var sharedStringPart = workbookPart.AddNewPart<SharedStringTablePart>();
                sharedStringPart.SharedStringTable = new SharedStringTable();

                var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                var sheetData = new SheetData(
                    CreateRow(workbookPart, 1, "Name", "Amount", "Status"),
                    CreateRow(workbookPart, 2, "{#lines}", string.Empty, string.Empty),
                    CreateRow(
                        3,
                        CreateSharedStringCell(workbookPart, "{name}", "A3"),
                        CreateSharedStringCell(workbookPart, "{amount}", "B3"),
                        CreateFormulaCell("B3*2", "C3")),
                    CreateRow(workbookPart, 4, "{/lines}", string.Empty, string.Empty),
                    CreateRow(
                        5,
                        CreateSharedStringCell(workbookPart, "Total", "A5"),
                        CreateSharedStringCell(workbookPart, string.Empty, "B5"),
                        CreateFormulaCell("SUM(C3:C3)", "C5")));

                worksheetPart.Worksheet = new Worksheet(
                    new SheetDimension { Reference = "A1:C5" },
                    new AutoFilter { Reference = "A1:C5" },
                    sheetData);

                var tableDefinitionPart = worksheetPart.AddNewPart<TableDefinitionPart>("rIdTable1");
                tableDefinitionPart.Table = new Table
                {
                    Id = 1U,
                    Name = "ReportTable",
                    DisplayName = "ReportTable",
                    Reference = "A1:C5",
                    TotalsRowShown = false
                };
                tableDefinitionPart.Table.Append(new AutoFilter { Reference = "A1:C5" });
                tableDefinitionPart.Table.Append(new TableColumns(
                    new TableColumn { Id = 1U, Name = "Name" },
                    new TableColumn { Id = 2U, Name = "Amount" },
                    new TableColumn { Id = 3U, Name = "Status" })
                { Count = 3U });
                tableDefinitionPart.Table.Save();

                worksheetPart.Worksheet.Append(new TableParts(
                    new TablePart { Id = worksheetPart.GetIdOfPart(tableDefinitionPart) })
                { Count = 1U });
                worksheetPart.Worksheet.Save();

                workbookPart.Workbook.Append(
                    new Sheets(
                        new Sheet
                        {
                            Id = workbookPart.GetIdOfPart(worksheetPart),
                            SheetId = 1U,
                            Name = "Report"
                        }),
                    new DefinedNames(
                        new DefinedName("'Report'!$A$1:$C$5")
                        {
                            Name = "ReportRange"
                        }));

                var calcChainPart = workbookPart.AddNewPart<CalculationChainPart>();
                calcChainPart.CalculationChain = new CalculationChain(new CalculationCell { CellReference = "C5", SheetId = 1 });
                calcChainPart.CalculationChain.Save();

                workbookPart.Workbook.Save();
            }

            return stream.ToArray();
        }
    }

    private static byte[] CreateWorkbookWithReferencedMediaCells(string? leadingCellReference, string? mediaCellReference)
    {
        using var stream = new MemoryStream();
        using (var document = SpreadsheetDocument.Create(stream, DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook, true))
        {
            var workbookPart = document.AddWorkbookPart();
            workbookPart.Workbook = new Workbook();

            var sharedStringPart = workbookPart.AddNewPart<SharedStringTablePart>();
            sharedStringPart.SharedStringTable = new SharedStringTable();

            var cells = new List<Cell>();
            if (leadingCellReference != null)
            {
                cells.Add(CreateSharedStringCell(workbookPart, "Label", leadingCellReference));
            }

            var mediaCell = CreateSharedStringCell(workbookPart, "{%image}", mediaCellReference ?? "A1");
            mediaCell.CellReference = mediaCellReference;
            cells.Add(mediaCell);

            var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = new Worksheet(
                new SheetDimension { Reference = "A1:D1" },
                new SheetData(CreateRow(1, cells.ToArray())));
            worksheetPart.Worksheet.Save();

            workbookPart.Workbook.Append(
                new Sheets(
                    new Sheet
                    {
                        Id = workbookPart.GetIdOfPart(worksheetPart),
                        SheetId = 1U,
                        Name = "Report"
                    }));
            workbookPart.Workbook.Save();
        }

        return stream.ToArray();
    }

    private static byte[] CreateWorkbookWithAutoFilterReference(string reference)
    {
        using var stream = new MemoryStream();
        using (var document = SpreadsheetDocument.Create(stream, DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook, true))
        {
            var workbookPart = document.AddWorkbookPart();
            workbookPart.Workbook = new Workbook();

            var sharedStringPart = workbookPart.AddNewPart<SharedStringTablePart>();
            sharedStringPart.SharedStringTable = new SharedStringTable();

            var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = new Worksheet(
                new SheetDimension { Reference = "A1" },
                new AutoFilter { Reference = reference },
                new SheetData(CreateRow(workbookPart, 1, "Value")));
            worksheetPart.Worksheet.Save();

            workbookPart.Workbook.Append(
                new Sheets(
                    new Sheet
                    {
                        Id = workbookPart.GetIdOfPart(worksheetPart),
                        SheetId = 1U,
                        Name = "Report"
                    }));
            workbookPart.Workbook.Save();
        }

        return stream.ToArray();
    }

    private static Row CreateRow(WorkbookPart workbookPart, int rowIndex, params string[] values)
    {
        return CreateRow(
            rowIndex,
            values.Select((value, index) => CreateSharedStringCell(workbookPart, value, GetCellReference(index + 1, rowIndex))).ToArray());
    }

    private static Row CreateRow(int rowIndex, params Cell[] cells)
    {
        var row = new Row { RowIndex = (uint)rowIndex };
        foreach (var cell in cells)
        {
            row.Append(cell);
        }

        return row;
    }

    private static Cell CreateSharedStringCell(WorkbookPart workbookPart, string text, string cellReference)
    {
        var sharedStringTable = workbookPart.SharedStringTablePart!.SharedStringTable!;
        sharedStringTable.AppendChild(new SharedStringItem(new Text(text ?? string.Empty)));
        var index = sharedStringTable.Elements<SharedStringItem>().Count() - 1;

        return new Cell
        {
            CellReference = cellReference,
            DataType = CellValues.SharedString,
            CellValue = new CellValue(index.ToString(CultureInfo.InvariantCulture))
        };
    }

    private static Cell CreateFormulaCell(string formula, string cellReference)
    {
        return new Cell
        {
            CellReference = cellReference,
            CellFormula = new CellFormula(formula),
            CellValue = new CellValue(string.Empty)
        };
    }

    private static Stylesheet CreateStylesheet()
    {
        return new Stylesheet(
            new Fonts(new Font(), new Font(new Bold())),
            new Fills(new Fill(new PatternFill { PatternType = PatternValues.None }), new Fill(new PatternFill { PatternType = PatternValues.Gray125 })),
            new Borders(new Border()),
            new CellStyleFormats(new CellFormat()),
            new CellFormats(
                new CellFormat(),
                new CellFormat { FontId = 1U, FillId = 0U, BorderId = 0U, ApplyFont = true }));
    }

    private static IReadOnlyList<SheetRowData> ReadSheetRows(byte[] workbookBytes)
    {
        using (var stream = new MemoryStream(workbookBytes))
        using (var document = SpreadsheetDocument.Open(stream, false))
        {
            var workbookPart = document.WorkbookPart!;
            var worksheetPart = workbookPart.WorksheetParts.First();
            var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>()!;

            return sheetData.Elements<Row>()
                .Select(row =>
                {
                    var cells = row.Elements<Cell>().ToArray();
                    return new SheetRowData(
                        cells.Select(cell => GetCellText(cell, workbookPart)).ToArray(),
                        cells.Select(cell => cell.StyleIndex?.Value).ToArray(),
                        cells.Select(cell => cell.DataType?.Value).ToArray());
                })
                .ToArray();
        }
    }

    private static string GetCellReference(int columnIndex, int rowIndex)
    {
        return GetColumnName(columnIndex) + rowIndex.ToString(CultureInfo.InvariantCulture);
    }

    private static string AnchorToCell(OneCellAnchor anchor)
    {
        var from = anchor.FromMarker!;
        var column = int.Parse(from.ColumnId!.Text!, CultureInfo.InvariantCulture) + 1;
        var row = int.Parse(from.RowId!.Text!, CultureInfo.InvariantCulture) + 1;
        return GetCellReference(column, row);
    }

    private static long PixelsToEmu(int pixels)
    {
        return pixels * 9525L;
    }

    private static string GetCellText(Cell cell, WorkbookPart workbookPart)
    {
        if (cell.CellFormula != null)
        {
            return cell.CellFormula.Text ?? string.Empty;
        }

        if (cell.DataType?.Value == CellValues.SharedString)
        {
            if (!int.TryParse(cell.CellValue?.InnerText, NumberStyles.Integer, CultureInfo.InvariantCulture, out var index))
            {
                return string.Empty;
            }

            var item = workbookPart.SharedStringTablePart?.SharedStringTable?.Elements<SharedStringItem>().ElementAtOrDefault(index);
            if (item == null)
            {
                return string.Empty;
            }

            if (item.Text != null)
            {
                return item.Text.Text ?? string.Empty;
            }

            return string.Concat(item.Descendants<Text>().Select(static text => text.Text));
        }

        if (cell.InlineString != null)
        {
            return string.Concat(cell.InlineString.Descendants<Text>().Select(static text => text.Text));
        }

        return cell.CellValue?.InnerText ?? string.Empty;
    }

    private static string GetColumnName(int columnIndex)
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

    private static bool ContainsDarkPixels(byte[] pngBytes)
    {
        using (var image = Image.Load<Rgba32>(pngBytes))
        {
            for (var y = 0; y < image.Height; y++)
            {
                for (var x = 0; x < image.Width; x++)
                {
                    var pixel = image[x, y];
                    if (pixel.A > 0 && (pixel.R < 200 || pixel.G < 200 || pixel.B < 200))
                    {
                        return true;
                    }
                }
            }
        }

        return false;
    }

    private static (int Width, int Height) GetImageSize(byte[] imageBytes)
    {
        using (var image = Image.Load<Rgba32>(imageBytes))
        {
            return (image.Width, image.Height);
        }
    }

    private readonly struct SheetRowData
    {
        public SheetRowData(string[] values, uint?[] styleIndexes, CellValues?[] dataTypes)
        {
            Values = values;
            StyleIndexes = styleIndexes;
            DataTypes = dataTypes;
        }

        public string[] Values { get; }

        public uint?[] StyleIndexes { get; }

        public CellValues?[] DataTypes { get; }
    }

    private readonly struct RowSpec
    {
        public RowSpec(CellSpec[] cells)
        {
            Cells = cells;
        }

        public CellSpec[] Cells { get; }

        public static RowSpec Create(params string[] values)
        {
            return new RowSpec(values.Select(static value => CellSpec.Text(value)).ToArray());
        }

        public static RowSpec Create(string[] values, uint styleIndex)
        {
            return new RowSpec(values.Select(value => CellSpec.Text(value, styleIndex)).ToArray());
        }

        public static RowSpec Create(params CellSpec[] cells)
        {
            return new RowSpec(cells);
        }
    }

    private enum CellKind
    {
        Text,
        Formula
    }

    private readonly struct CellSpec
    {
        private CellSpec(string value, CellKind kind, uint? styleIndex)
        {
            Value = value;
            Kind = kind;
            StyleIndex = styleIndex;
        }

        public string Value { get; }

        public CellKind Kind { get; }

        public uint? StyleIndex { get; }

        public static CellSpec Text(string value, uint? styleIndex = null)
        {
            return new CellSpec(value, CellKind.Text, styleIndex);
        }

        public static CellSpec Formula(string formula, uint? styleIndex = null)
        {
            return new CellSpec(formula, CellKind.Formula, styleIndex);
        }
    }
}
