using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.Json.Nodes;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using JToken = System.Text.Json.Nodes.JsonNode;

namespace NDocxTemplater;

public sealed class XlsxTemplateEngine
{
    public byte[] Render(byte[] templateBytes, string jsonData)
    {
        if (templateBytes == null)
        {
            throw new ArgumentNullException(nameof(templateBytes));
        }

        using (var templateStream = new MemoryStream(templateBytes, writable: false))
        using (var outputStream = new MemoryStream())
        {
            Render(templateStream, outputStream, jsonData);
            return outputStream.ToArray();
        }
    }

    public void Render(Stream templateStream, Stream outputStream, string jsonData)
    {
        if (templateStream == null)
        {
            throw new ArgumentNullException(nameof(templateStream));
        }

        if (outputStream == null)
        {
            throw new ArgumentNullException(nameof(outputStream));
        }

        if (jsonData == null)
        {
            throw new ArgumentNullException(nameof(jsonData));
        }

        if (!outputStream.CanSeek || !outputStream.CanWrite)
        {
            throw new ArgumentException("Output stream must be seekable and writable.", nameof(outputStream));
        }

        outputStream.SetLength(0);
        templateStream.Position = 0;
        templateStream.CopyTo(outputStream);
        outputStream.Position = 0;

        var rootData = JsonNode.Parse(jsonData);
        if (rootData == null)
        {
            throw new InvalidOperationException("The JSON data could not be parsed.");
        }

        using (var document = SpreadsheetDocument.Open(outputStream, true))
        {
            if (document.WorkbookPart?.Workbook == null)
            {
                throw new InvalidOperationException("The XLSX template does not contain a valid workbook.");
            }

            var renderer = new SpreadsheetTemplateRenderer(document.WorkbookPart, rootData);
            renderer.Render();
            document.WorkbookPart.Workbook.Save();
        }

        outputStream.Position = 0;
    }
}

internal sealed class SpreadsheetTemplateRenderer
{
    private readonly WorkbookPart _workbookPart;
    private readonly JToken _rootData;

    public SpreadsheetTemplateRenderer(WorkbookPart workbookPart, JToken rootData)
    {
        _workbookPart = workbookPart;
        _rootData = rootData;
    }

    public void Render()
    {
        var rootContext = new TemplateContext(_rootData, _rootData, null);
        foreach (var worksheetPart in _workbookPart.WorksheetParts)
        {
            RenderWorksheet(worksheetPart, rootContext);
        }
    }

    private void RenderWorksheet(WorksheetPart worksheetPart, TemplateContext context)
    {
        var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();
        if (sheetData == null)
        {
            return;
        }

        var sourceRows = sheetData.Elements<Row>().ToList();
        var renderedRows = new List<Row>();

        for (var index = 0; index < sourceRows.Count; index++)
        {
            var marker = SpreadsheetControlMarker.TryParse(sourceRows[index], _workbookPart);

            if (marker != null && marker.IsStart)
            {
                var endIndex = FindMatchingEnd(sourceRows, index, marker, _workbookPart);
                var blockTemplates = sourceRows.Skip(index + 1).Take(endIndex - index - 1).ToList();

                if (marker.Kind == ControlMarkerKind.LoopStart)
                {
                    var loopData = ExpressionEvaluator.Evaluate(marker.Expression, context);
                    foreach (var item in ExpressionEvaluator.ToLoopItems(loopData))
                    {
                        var itemContext = new TemplateContext(item, _rootData, context);
                        RenderRows(blockTemplates, renderedRows, itemContext);
                    }
                }
                else if (marker.Kind == ControlMarkerKind.IfStart)
                {
                    var conditionValue = ExpressionEvaluator.Evaluate(marker.Expression, context);
                    if (ExpressionEvaluator.IsTruthy(conditionValue))
                    {
                        RenderRows(blockTemplates, renderedRows, context);
                    }
                }

                index = endIndex;
                continue;
            }

            if (marker != null && marker.IsEnd)
            {
                continue;
            }

            var cloned = (Row)sourceRows[index].CloneNode(true);
            RenderRow(cloned, context);
            renderedRows.Add(cloned);
        }

        sheetData.RemoveAllChildren<Row>();
        foreach (var row in renderedRows)
        {
            sheetData.Append(row);
        }

        SpreadsheetCellHelper.ReindexRows(sheetData);
        SpreadsheetCellHelper.UpdateSheetDimension(worksheetPart.Worksheet, sheetData);
        worksheetPart.Worksheet.Save();
    }

    private void RenderRows(IReadOnlyCollection<Row> templates, ICollection<Row> renderedRows, TemplateContext context)
    {
        foreach (var template in templates)
        {
            var clone = (Row)template.CloneNode(true);
            RenderRow(clone, context);
            renderedRows.Add(clone);
        }
    }

    private void RenderRow(Row row, TemplateContext context)
    {
        foreach (var cell in row.Elements<Cell>())
        {
            RenderCell(cell, context);
        }
    }

    private void RenderCell(Cell cell, TemplateContext context)
    {
        var originalText = SpreadsheetCellHelper.GetCellText(cell, _workbookPart);
        if (string.IsNullOrEmpty(originalText))
        {
            return;
        }

        if (originalText.IndexOf('{') < 0 || originalText.IndexOf('}') < 0)
        {
            return;
        }

        var trimmed = originalText.Trim();
        var fullTagMatch = TagPatterns.SingleTagRegex.Match(trimmed);
        if (fullTagMatch.Success)
        {
            var expression = fullTagMatch.Groups[1].Value.Trim();
            if (ControlMarker.IsControlToken(expression))
            {
                SpreadsheetCellHelper.SetCellString(cell, string.Empty);
                return;
            }

            SpreadsheetCellHelper.SetCellValue(cell, ExpressionEvaluator.Evaluate(expression, context));
            return;
        }

        var replaced = TagPatterns.InlineTagRegex.Replace(originalText, match =>
        {
            var expression = match.Groups[1].Value.Trim();
            if (ControlMarker.IsControlToken(expression))
            {
                return string.Empty;
            }

            return ExpressionEvaluator.ToText(ExpressionEvaluator.Evaluate(expression, context));
        });

        SpreadsheetCellHelper.SetCellString(cell, replaced);
    }

    private static int FindMatchingEnd(
        IReadOnlyList<Row> rows,
        int startIndex,
        SpreadsheetControlMarker startMarker,
        WorkbookPart workbookPart)
    {
        var depth = 0;

        for (var index = startIndex + 1; index < rows.Count; index++)
        {
            var marker = SpreadsheetControlMarker.TryParse(rows[index], workbookPart);
            if (marker == null)
            {
                continue;
            }

            if (ControlMarker.IsStartOfSameType(startMarker.Kind, marker.Kind))
            {
                depth++;
                continue;
            }

            if (!ControlMarker.IsEndOfSameType(startMarker.Kind, marker.Kind))
            {
                continue;
            }

            if (depth > 0)
            {
                depth--;
                continue;
            }

            if (!string.Equals(marker.Expression, startMarker.Expression, StringComparison.Ordinal))
            {
                throw new InvalidOperationException(
                    string.Format(
                        CultureInfo.InvariantCulture,
                        "Closing tag '{0}' does not match opening tag '{1}'.",
                        marker.RawToken,
                        startMarker.RawToken));
            }

            return index;
        }

        throw new InvalidOperationException(
            string.Format(
                CultureInfo.InvariantCulture,
                "No closing tag found for '{0}'.",
                startMarker.RawToken));
    }
}

internal sealed class SpreadsheetControlMarker
{
    private SpreadsheetControlMarker(ControlMarkerKind kind, string expression, string rawToken)
    {
        Kind = kind;
        Expression = expression;
        RawToken = rawToken;
    }

    public ControlMarkerKind Kind { get; }

    public string Expression { get; }

    public string RawToken { get; }

    public bool IsStart => Kind == ControlMarkerKind.LoopStart || Kind == ControlMarkerKind.IfStart;

    public bool IsEnd => Kind == ControlMarkerKind.LoopEnd || Kind == ControlMarkerKind.IfEnd;

    public static SpreadsheetControlMarker? TryParse(Row row, WorkbookPart? workbookPart)
    {
        var nonEmptyTexts = row.Elements<Cell>()
            .Select(cell => SpreadsheetCellHelper.GetCellText(cell, workbookPart))
            .Select(static text => text.Trim())
            .Where(static text => text.Length > 0)
            .ToArray();

        if (nonEmptyTexts.Length != 1)
        {
            return null;
        }

        var fullTagMatch = TagPatterns.SingleTagRegex.Match(nonEmptyTexts[0]);
        if (!fullTagMatch.Success)
        {
            return null;
        }

        var token = fullTagMatch.Groups[1].Value.Trim();
        if (token.Length == 0)
        {
            return null;
        }

        if (token.StartsWith("#", StringComparison.Ordinal))
        {
            var expression = token.Substring(1).Trim();
            return expression.Length == 0 ? null : new SpreadsheetControlMarker(ControlMarkerKind.LoopStart, expression, "{" + token + "}");
        }

        if (token.StartsWith("/?", StringComparison.Ordinal))
        {
            var expression = token.Substring(2).Trim();
            return expression.Length == 0 ? null : new SpreadsheetControlMarker(ControlMarkerKind.IfEnd, expression, "{" + token + "}");
        }

        if (token.StartsWith("?", StringComparison.Ordinal))
        {
            var expression = token.Substring(1).Trim();
            return expression.Length == 0 ? null : new SpreadsheetControlMarker(ControlMarkerKind.IfStart, expression, "{" + token + "}");
        }

        if (token.StartsWith("/", StringComparison.Ordinal))
        {
            var expression = token.Substring(1).Trim();
            return expression.Length == 0 ? null : new SpreadsheetControlMarker(ControlMarkerKind.LoopEnd, expression, "{" + token + "}");
        }

        return null;
    }
}

internal static class SpreadsheetCellHelper
{
    public static string GetCellText(Cell cell, WorkbookPart? workbookPart)
    {
        if (cell == null)
        {
            return string.Empty;
        }

        if (cell.CellFormula != null)
        {
            return cell.CellFormula.Text ?? string.Empty;
        }

        if (cell.DataType?.Value == CellValues.SharedString)
        {
            if (workbookPart?.SharedStringTablePart?.SharedStringTable == null)
            {
                return string.Empty;
            }

            if (!int.TryParse(cell.CellValue?.InnerText, NumberStyles.Integer, CultureInfo.InvariantCulture, out var sharedStringIndex))
            {
                return string.Empty;
            }

            var item = workbookPart.SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>().ElementAtOrDefault(sharedStringIndex);
            return item == null ? string.Empty : GetSharedStringText(item);
        }

        if (cell.InlineString != null)
        {
            return string.Concat(cell.InlineString.Descendants<Text>().Select(static text => text.Text));
        }

        return cell.CellValue?.InnerText ?? string.Empty;
    }

    public static void SetCellValue(Cell cell, JToken? token)
    {
        if (JsonNodeHelpers.IsNull(token))
        {
            SetCellString(cell, string.Empty);
            return;
        }

        if (JsonNodeHelpers.TryGetBoolean(token, out var boolValue))
        {
            ResetCellContent(cell);
            cell.DataType = CellValues.Boolean;
            cell.CellValue = new CellValue(boolValue ? "1" : "0");
            return;
        }

        if (JsonNodeHelpers.TryGetDecimal(token, out var decimalValue))
        {
            ResetCellContent(cell);
            cell.DataType = null;
            cell.CellValue = new CellValue(decimalValue.ToString(CultureInfo.InvariantCulture));
            return;
        }

        SetCellString(cell, ExpressionEvaluator.ToText(token));
    }

    public static void SetCellString(Cell cell, string text)
    {
        ResetCellContent(cell);
        cell.DataType = CellValues.InlineString;
        cell.InlineString = new InlineString(CreateTextNode(text));
    }

    public static void ReindexRows(SheetData sheetData)
    {
        uint rowIndex = 1;
        foreach (var row in sheetData.Elements<Row>())
        {
            row.RowIndex = rowIndex;

            var cellIndex = 1;
            foreach (var cell in row.Elements<Cell>())
            {
                var columnName = GetColumnName(cell, cellIndex);
                cell.CellReference = columnName + rowIndex.ToString(CultureInfo.InvariantCulture);
                cellIndex++;
            }

            row.Spans = null;
            rowIndex++;
        }
    }

    public static void UpdateSheetDimension(Worksheet worksheet, SheetData sheetData)
    {
        var rows = sheetData.Elements<Row>().ToList();
        var dimensionReference = "A1";

        if (rows.Count > 0)
        {
            var maxRow = rows.Max(static row => row.RowIndex?.Value ?? 1U);
            var maxColumnIndex = 1;

            foreach (var row in rows)
            {
                foreach (var cell in row.Elements<Cell>())
                {
                    maxColumnIndex = Math.Max(maxColumnIndex, GetColumnIndex(cell.CellReference?.Value));
                }
            }

            dimensionReference = "A1:" + GetColumnName(maxColumnIndex) + maxRow.ToString(CultureInfo.InvariantCulture);
        }

        var dimension = worksheet.GetFirstChild<SheetDimension>();
        if (dimension == null)
        {
            var insertIndex = worksheet.GetFirstChild<SheetProperties>() == null ? 0 : 1;
            dimension = worksheet.InsertAt(new SheetDimension(), insertIndex);
        }

        dimension.Reference = dimensionReference;
    }

    private static Text CreateTextNode(string text)
    {
        var safeText = text ?? string.Empty;
        var node = new Text(safeText);
        if (safeText.Length > 0 && (safeText.StartsWith(" ", StringComparison.Ordinal) || safeText.EndsWith(" ", StringComparison.Ordinal)))
        {
            node.Space = SpaceProcessingModeValues.Preserve;
        }

        return node;
    }

    private static void ResetCellContent(Cell cell)
    {
        cell.CellFormula = null;
        cell.CellValue = null;
        cell.InlineString = null;
        cell.RemoveAllChildren<CellFormula>();
        cell.RemoveAllChildren<CellValue>();
        cell.RemoveAllChildren<InlineString>();
    }

    private static string GetSharedStringText(SharedStringItem item)
    {
        if (item.Text != null)
        {
            return item.Text.Text ?? string.Empty;
        }

        return string.Concat(item.Descendants<Text>().Select(static text => text.Text));
    }

    private static string GetColumnName(Cell cell, int fallbackIndex)
    {
        var reference = cell.CellReference?.Value;
        var letters = new string((reference ?? string.Empty).TakeWhile(static ch => !char.IsDigit(ch)).ToArray());
        return letters.Length > 0 ? letters : GetColumnName(fallbackIndex);
    }

    private static int GetColumnIndex(string? cellReference)
    {
        var letters = new string((cellReference ?? string.Empty).TakeWhile(static ch => !char.IsDigit(ch)).ToArray());
        if (letters.Length == 0)
        {
            return 1;
        }

        var columnIndex = 0;
        foreach (var ch in letters)
        {
            columnIndex = (columnIndex * 26) + (char.ToUpperInvariant(ch) - 'A' + 1);
        }

        return columnIndex;
    }

    private static string GetColumnName(int columnIndex)
    {
        if (columnIndex <= 0)
        {
            return "A";
        }

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
}
