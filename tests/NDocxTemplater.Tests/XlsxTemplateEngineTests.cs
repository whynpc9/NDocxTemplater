using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Xunit;

namespace NDocxTemplater.Tests;

public class XlsxTemplateEngineTests
{
    private readonly XlsxTemplateEngine _engine = new XlsxTemplateEngine();

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

    private static byte[] CreateWorkbook(params RowSpec[] rows)
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
                    for (var columnIndex = 0; columnIndex < rows[rowIndex].Values.Length; columnIndex++)
                    {
                        var cell = CreateSharedStringCell(
                            workbookPart,
                            rows[rowIndex].Values[columnIndex],
                            GetCellReference(columnIndex + 1, rowIndex + 1));

                        var styleIndex = rows[rowIndex].GetStyleIndex(columnIndex);
                        if (styleIndex.HasValue)
                        {
                            cell.StyleIndex = styleIndex.Value;
                        }

                        row.Append(cell);
                    }

                    sheetData.Append(row);
                }

                worksheetPart.Worksheet = new Worksheet(
                    new SheetDimension { Reference = "A1:" + GetCellReference(rows[0].Values.Length, rows.Length) },
                    sheetData);
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

    private static string GetCellText(Cell cell, WorkbookPart workbookPart)
    {
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

    private readonly struct RowSpec
    {
        private RowSpec(string[] values, uint?[] styleIndexes)
        {
            Values = values;
            StyleIndexes = styleIndexes;
        }

        public string[] Values { get; }

        public uint?[] StyleIndexes { get; }

        public static RowSpec Create(params string[] values)
        {
            return new RowSpec(values, new uint?[values.Length]);
        }

        public static RowSpec Create(string[] values, uint styleIndex)
        {
            return new RowSpec(values, Enumerable.Repeat<uint?>(styleIndex, values.Length).ToArray());
        }

        public uint? GetStyleIndex(int columnIndex)
        {
            return columnIndex >= 0 && columnIndex < StyleIndexes.Length ? StyleIndexes[columnIndex] : null;
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
}
