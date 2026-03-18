using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.Json.Nodes;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;
using S = DocumentFormat.OpenXml.Spreadsheet;
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
    private const string RootScopeId = "$root";

    private readonly WorkbookPart _workbookPart;
    private readonly JToken _rootData;
    private uint _drawingObjectIdCounter;
    private int _scopeCounter;

    public SpreadsheetTemplateRenderer(WorkbookPart workbookPart, JToken rootData)
    {
        _workbookPart = workbookPart;
        _rootData = rootData;
        _drawingObjectIdCounter = SpreadsheetDrawingHelper.GetNextObjectId(workbookPart);
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
        var sheetData = worksheetPart.Worksheet.GetFirstChild<S.SheetData>();
        if (sheetData == null)
        {
            return;
        }

        var sourceRows = sheetData.Elements<S.Row>()
            .Select((row, index) => new SpreadsheetSourceRow(row, row.RowIndex?.Value ?? (uint)(index + 1)))
            .ToList();
        var originalMergeReferences = SpreadsheetMergeHelper.ReadMergeReferences(worksheetPart.Worksheet);
        var renderedRows = new List<RenderedSpreadsheetRow>();

        RenderRows(sourceRows, renderedRows, context, RootScopeId);

        for (var index = 0; index < renderedRows.Count; index++)
        {
            var targetRowIndex = (uint)(index + 1);
            renderedRows[index].TargetRowIndex = targetRowIndex;
            SpreadsheetCellHelper.UpdateRowIndex(renderedRows[index].Row, targetRowIndex);
        }

        var rowMapping = new SpreadsheetRowMapping(renderedRows, RootScopeId);
        foreach (var renderedRow in renderedRows)
        {
            SpreadsheetFormulaHelper.RewriteRowFormulas(renderedRow.Row, renderedRow.ScopeId, rowMapping);
        }

        sheetData.RemoveAllChildren<S.Row>();
        foreach (var renderedRow in renderedRows)
        {
            sheetData.Append(renderedRow.Row);
        }

        SpreadsheetMergeHelper.RebuildMergeCells(worksheetPart.Worksheet, originalMergeReferences, rowMapping);
        SpreadsheetDrawingHelper.RenderMedia(worksheetPart, renderedRows, NextDrawingObjectId);
        SpreadsheetCellHelper.UpdateSheetDimension(worksheetPart.Worksheet, sheetData);
        worksheetPart.Worksheet.Save();
    }

    private void RenderRows(
        IReadOnlyList<SpreadsheetSourceRow> templates,
        ICollection<RenderedSpreadsheetRow> renderedRows,
        TemplateContext context,
        string scopeId)
    {
        for (var index = 0; index < templates.Count; index++)
        {
            var sourceRow = templates[index];
            var marker = SpreadsheetControlMarker.TryParse(sourceRow.Row, _workbookPart);

            if (marker != null && marker.IsStart)
            {
                var endIndex = FindMatchingEnd(templates, index, marker, _workbookPart);
                var blockTemplates = templates.Skip(index + 1).Take(endIndex - index - 1).ToList();

                if (marker.Kind == ControlMarkerKind.LoopStart)
                {
                    var loopData = ExpressionEvaluator.Evaluate(marker.Expression, context);
                    foreach (var item in ExpressionEvaluator.ToLoopItems(loopData))
                    {
                        var itemContext = new TemplateContext(item, _rootData, context);
                        RenderRows(blockTemplates, renderedRows, itemContext, CreateChildScopeId(scopeId));
                    }
                }
                else if (marker.Kind == ControlMarkerKind.IfStart)
                {
                    var conditionValue = ExpressionEvaluator.Evaluate(marker.Expression, context);
                    if (ExpressionEvaluator.IsTruthy(conditionValue))
                    {
                        RenderRows(blockTemplates, renderedRows, context, scopeId);
                    }
                }

                index = endIndex;
                continue;
            }

            if (marker != null && marker.IsEnd)
            {
                continue;
            }

            var clone = (S.Row)sourceRow.Row.CloneNode(true);
            var renderedRow = new RenderedSpreadsheetRow(clone, sourceRow.SourceRowIndex, scopeId);
            RenderRow(renderedRow, context);
            renderedRows.Add(renderedRow);
        }
    }

    private void RenderRow(RenderedSpreadsheetRow row, TemplateContext context)
    {
        var columnIndex = 1;
        foreach (var cell in row.Row.Elements<S.Cell>())
        {
            RenderCell(cell, columnIndex, row, context);
            columnIndex++;
        }
    }

    private void RenderCell(S.Cell cell, int columnIndex, RenderedSpreadsheetRow row, TemplateContext context)
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

            if (ImageTagParser.TryParseToken(expression, out var imageTag))
            {
                var payloads = TemplateMediaResolver.ResolveMany(imageTag.Expression, context).ToList();
                if (payloads.Count > 0)
                {
                    row.MediaPlacements.Add(new SpreadsheetMediaPlacement(columnIndex, payloads));
                }

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

            if (ImageTagParser.TryParseToken(expression, out _))
            {
                return string.Empty;
            }

            return ExpressionEvaluator.ToText(ExpressionEvaluator.Evaluate(expression, context));
        });

        SpreadsheetCellHelper.SetCellString(cell, replaced);
    }

    private string CreateChildScopeId(string parentScopeId)
    {
        _scopeCounter++;
        return parentScopeId + "/" + _scopeCounter.ToString(CultureInfo.InvariantCulture);
    }

    private uint NextDrawingObjectId()
    {
        return _drawingObjectIdCounter++;
    }

    private static int FindMatchingEnd(
        IReadOnlyList<SpreadsheetSourceRow> rows,
        int startIndex,
        SpreadsheetControlMarker startMarker,
        WorkbookPart workbookPart)
    {
        var depth = 0;

        for (var index = startIndex + 1; index < rows.Count; index++)
        {
            var marker = SpreadsheetControlMarker.TryParse(rows[index].Row, workbookPart);
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

    public static SpreadsheetControlMarker? TryParse(S.Row row, WorkbookPart? workbookPart)
    {
        var nonEmptyTexts = row.Elements<S.Cell>()
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
    public static string GetCellText(S.Cell cell, WorkbookPart? workbookPart)
    {
        if (cell == null)
        {
            return string.Empty;
        }

        if (cell.CellFormula != null)
        {
            return cell.CellFormula.Text ?? string.Empty;
        }

        if (cell.DataType?.Value == S.CellValues.SharedString)
        {
            if (workbookPart?.SharedStringTablePart?.SharedStringTable == null)
            {
                return string.Empty;
            }

            if (!int.TryParse(cell.CellValue?.InnerText, NumberStyles.Integer, CultureInfo.InvariantCulture, out var sharedStringIndex))
            {
                return string.Empty;
            }

            var item = workbookPart.SharedStringTablePart.SharedStringTable.Elements<S.SharedStringItem>().ElementAtOrDefault(sharedStringIndex);
            return item == null ? string.Empty : GetSharedStringText(item);
        }

        if (cell.InlineString != null)
        {
            return string.Concat(cell.InlineString.Descendants<S.Text>().Select(static text => text.Text));
        }

        return cell.CellValue?.InnerText ?? string.Empty;
    }

    public static void SetCellValue(S.Cell cell, JToken? token)
    {
        if (JsonNodeHelpers.IsNull(token))
        {
            SetCellString(cell, string.Empty);
            return;
        }

        if (JsonNodeHelpers.TryGetBoolean(token, out var boolValue))
        {
            ResetCellContent(cell);
            cell.DataType = S.CellValues.Boolean;
            cell.CellValue = new S.CellValue(boolValue ? "1" : "0");
            return;
        }

        if (JsonNodeHelpers.TryGetDecimal(token, out var decimalValue))
        {
            ResetCellContent(cell);
            cell.DataType = null;
            cell.CellValue = new S.CellValue(decimalValue.ToString(CultureInfo.InvariantCulture));
            return;
        }

        SetCellString(cell, ExpressionEvaluator.ToText(token));
    }

    public static void SetCellString(S.Cell cell, string text)
    {
        ResetCellContent(cell);
        cell.DataType = S.CellValues.InlineString;
        cell.InlineString = new S.InlineString(CreateTextNode(text));
    }

    public static void UpdateRowIndex(S.Row row, uint rowIndex)
    {
        row.RowIndex = rowIndex;

        var cellIndex = 1;
        foreach (var cell in row.Elements<S.Cell>())
        {
            var columnName = GetColumnName(cell, cellIndex);
            cell.CellReference = columnName + rowIndex.ToString(CultureInfo.InvariantCulture);
            cellIndex++;
        }

        row.Spans = null;
    }

    public static void UpdateSheetDimension(S.Worksheet worksheet, S.SheetData sheetData)
    {
        var rows = sheetData.Elements<S.Row>().ToList();
        var dimensionReference = "A1";

        if (rows.Count > 0)
        {
            var maxRow = rows.Max(static row => row.RowIndex?.Value ?? 1U);
            var maxColumnIndex = 1;

            foreach (var row in rows)
            {
                foreach (var cell in row.Elements<S.Cell>())
                {
                    maxColumnIndex = Math.Max(maxColumnIndex, GetColumnIndex(cell.CellReference?.Value));
                }
            }

            dimensionReference = "A1:" + GetColumnName(maxColumnIndex) + maxRow.ToString(CultureInfo.InvariantCulture);
        }

        var dimension = worksheet.GetFirstChild<S.SheetDimension>();
        if (dimension == null)
        {
            var insertIndex = worksheet.GetFirstChild<S.SheetProperties>() == null ? 0 : 1;
            dimension = worksheet.InsertAt(new S.SheetDimension(), insertIndex);
        }

        dimension.Reference = dimensionReference;
    }

    public static int GetColumnIndex(string? cellReference)
    {
        var letters = new string((cellReference ?? string.Empty).TakeWhile(static ch => !char.IsDigit(ch)).ToArray());
        return GetColumnIndexFromLetters(letters);
    }

    public static string GetColumnName(int columnIndex)
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

    private static int GetColumnIndexFromLetters(string letters)
    {
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

    private static S.Text CreateTextNode(string text)
    {
        var safeText = text ?? string.Empty;
        var node = new S.Text(safeText);
        if (safeText.Length > 0 && (safeText.StartsWith(" ", StringComparison.Ordinal) || safeText.EndsWith(" ", StringComparison.Ordinal)))
        {
            node.Space = SpaceProcessingModeValues.Preserve;
        }

        return node;
    }

    private static void ResetCellContent(S.Cell cell)
    {
        cell.CellFormula = null;
        cell.CellValue = null;
        cell.InlineString = null;
        cell.RemoveAllChildren<S.CellFormula>();
        cell.RemoveAllChildren<S.CellValue>();
        cell.RemoveAllChildren<S.InlineString>();
    }

    private static string GetSharedStringText(S.SharedStringItem item)
    {
        if (item.Text != null)
        {
            return item.Text.Text ?? string.Empty;
        }

        return string.Concat(item.Descendants<S.Text>().Select(static text => text.Text));
    }

    private static string GetColumnName(S.Cell cell, int fallbackIndex)
    {
        var reference = cell.CellReference?.Value;
        var letters = new string((reference ?? string.Empty).TakeWhile(static ch => !char.IsDigit(ch)).ToArray());
        return letters.Length > 0 ? letters : GetColumnName(fallbackIndex);
    }
}

internal static class SpreadsheetFormulaHelper
{
    private static readonly Regex RangeRegex = new Regex(
        @"(?<start>(?:(?:'[^']+'|[A-Za-z_][A-Za-z0-9_.]*)!)?\$?[A-Z]{1,3}\$?\d+):(?<end>(?:(?:'[^']+'|[A-Za-z_][A-Za-z0-9_.]*)!)?\$?[A-Z]{1,3}\$?\d+)",
        RegexOptions.Compiled);

    private static readonly Regex CellReferenceRegex = new Regex(
        @"(?:(?:'[^']+'|[A-Za-z_][A-Za-z0-9_.]*)!)?\$?[A-Z]{1,3}\$?\d+",
        RegexOptions.Compiled);

    public static void RewriteRowFormulas(S.Row row, string scopeId, SpreadsheetRowMapping mapping)
    {
        foreach (var cell in row.Elements<S.Cell>())
        {
            if (cell.CellFormula == null || string.IsNullOrWhiteSpace(cell.CellFormula.Text))
            {
                continue;
            }

            cell.CellFormula.Text = RewriteFormula(cell.CellFormula.Text!, scopeId, mapping);
        }
    }

    private static string RewriteFormula(string formula, string scopeId, SpreadsheetRowMapping mapping)
    {
        var preservedRanges = new List<string>();
        var rangeExpanded = RangeRegex.Replace(formula, match =>
        {
            var replacement = RewriteRange(match.Groups["start"].Value, match.Groups["end"].Value, scopeId, mapping);
            preservedRanges.Add(replacement);
            return "__NDT_RANGE_" + (preservedRanges.Count - 1).ToString(CultureInfo.InvariantCulture) + "__";
        });

        var rewritten = CellReferenceRegex.Replace(rangeExpanded, match => RewriteSingleReference(match.Value, scopeId, mapping));

        for (var index = 0; index < preservedRanges.Count; index++)
        {
            rewritten = rewritten.Replace(
                "__NDT_RANGE_" + index.ToString(CultureInfo.InvariantCulture) + "__",
                preservedRanges[index]);
        }

        return rewritten;
    }

    private static string RewriteRange(string startReference, string endReference, string scopeId, SpreadsheetRowMapping mapping)
    {
        if (!SpreadsheetCellReference.TryParse(startReference, out var start) || !SpreadsheetCellReference.TryParse(endReference, out var end))
        {
            return startReference + ":" + endReference;
        }

        if (start.RowAbsolute || end.RowAbsolute)
        {
            var rewrittenStart = RewriteSingleReference(startReference, scopeId, mapping);
            var rewrittenEnd = RewriteSingleReference(endReference, scopeId, mapping);
            return rewrittenStart + ":" + rewrittenEnd;
        }

        if (mapping.TryResolveRange(scopeId, start.RowIndex, end.RowIndex, out var mappedStartRow, out var mappedEndRow))
        {
            return start.WithRow(mappedStartRow) + ":" + end.WithRow(mappedEndRow);
        }

        return RewriteSingleReference(startReference, scopeId, mapping)
            + ":"
            + RewriteSingleReference(endReference, scopeId, mapping);
    }

    private static string RewriteSingleReference(string reference, string scopeId, SpreadsheetRowMapping mapping)
    {
        if (!SpreadsheetCellReference.TryParse(reference, out var cellReference))
        {
            return reference;
        }

        if (cellReference.RowAbsolute)
        {
            return reference;
        }

        if (mapping.TryResolveRow(scopeId, cellReference.RowIndex, out var targetRow))
        {
            return cellReference.WithRow(targetRow);
        }

        return reference;
    }
}

internal static class SpreadsheetMergeHelper
{
    public static IReadOnlyList<string> ReadMergeReferences(S.Worksheet worksheet)
    {
        var mergeCells = worksheet.GetFirstChild<S.MergeCells>();
        if (mergeCells == null)
        {
            return Array.Empty<string>();
        }

        return mergeCells.Elements<S.MergeCell>()
            .Select(static mergeCell => mergeCell.Reference?.Value)
            .Where(static reference => !string.IsNullOrWhiteSpace(reference))
            .Cast<string>()
            .ToList();
    }

    public static void RebuildMergeCells(S.Worksheet worksheet, IReadOnlyList<string> sourceMergeReferences, SpreadsheetRowMapping mapping)
    {
        var existingMergeCells = worksheet.GetFirstChild<S.MergeCells>();
        if (existingMergeCells != null)
        {
            existingMergeCells.Remove();
        }

        if (sourceMergeReferences.Count == 0)
        {
            return;
        }

        var mergedReferences = new List<string>();
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        foreach (var mergeReference in sourceMergeReferences)
        {
            if (!SpreadsheetRangeReference.TryParse(mergeReference, out var range))
            {
                if (seen.Add(mergeReference))
                {
                    mergedReferences.Add(mergeReference);
                }

                continue;
            }

            var addedAny = false;
            foreach (var scopeId in mapping.ScopeIds)
            {
                if (!mapping.TryResolveRangeInScope(scopeId, range.Start.RowIndex, range.End.RowIndex, out var mappedStartRow, out var mappedEndRow))
                {
                    continue;
                }

                var rewritten = range.WithRows(mappedStartRow, mappedEndRow);
                if (seen.Add(rewritten))
                {
                    mergedReferences.Add(rewritten);
                }

                addedAny = true;
            }

            if (addedAny || !mapping.TryResolveRangeGlobal(range.Start.RowIndex, range.End.RowIndex, out var globalStartRow, out var globalEndRow))
            {
                continue;
            }

            var globalReference = range.WithRows(globalStartRow, globalEndRow);
            if (seen.Add(globalReference))
            {
                mergedReferences.Add(globalReference);
            }
        }

        if (mergedReferences.Count == 0)
        {
            return;
        }

        var mergeCells = new S.MergeCells();
        foreach (var mergeReference in mergedReferences)
        {
            mergeCells.Append(new S.MergeCell { Reference = mergeReference });
        }

        mergeCells.Count = (uint)mergedReferences.Count;

        var customSheetView = worksheet.Elements<S.CustomSheetView>().LastOrDefault();
        if (customSheetView != null)
        {
            worksheet.InsertAfter(mergeCells, customSheetView);
            return;
        }

        var sheetData = worksheet.GetFirstChild<S.SheetData>();
        if (sheetData != null)
        {
            worksheet.InsertAfter(mergeCells, sheetData);
            return;
        }

        worksheet.Append(mergeCells);
    }
}

internal static class SpreadsheetDrawingHelper
{
    private const long EmusPerPixel = 9525L;

    public static uint GetNextObjectId(WorkbookPart workbookPart)
    {
        uint maxId = 0;

        foreach (var worksheetPart in workbookPart.WorksheetParts)
        {
            var drawingPart = worksheetPart.DrawingsPart;
            if (drawingPart?.WorksheetDrawing == null)
            {
                continue;
            }

            foreach (var property in drawingPart.WorksheetDrawing.Descendants<Xdr.NonVisualDrawingProperties>())
            {
                maxId = Math.Max(maxId, property.Id?.Value ?? 0U);
            }
        }

        return maxId + 1U;
    }

    public static void RenderMedia(WorksheetPart worksheetPart, IReadOnlyList<RenderedSpreadsheetRow> rows, Func<uint> nextObjectId)
    {
        var placements = rows
            .Where(static row => row.MediaPlacements.Count > 0)
            .ToList();

        if (placements.Count == 0)
        {
            return;
        }

        var drawingsPart = EnsureDrawingsPart(worksheetPart);
        var worksheetDrawing = drawingsPart.WorksheetDrawing ?? new Xdr.WorksheetDrawing();
        if (drawingsPart.WorksheetDrawing == null)
        {
            drawingsPart.WorksheetDrawing = worksheetDrawing;
        }

        foreach (var row in placements)
        {
            foreach (var placement in row.MediaPlacements)
            {
                var rowOffsetEmu = 0L;
                foreach (var payload in placement.Payloads)
                {
                    var imagePart = drawingsPart.AddImagePart(payload.ImagePartType);
                    using (var stream = new MemoryStream(payload.Bytes, writable: false))
                    {
                        imagePart.FeedData(stream);
                    }

                    var relationId = drawingsPart.GetIdOfPart(imagePart);
                    worksheetDrawing.Append(CreateAnchor(
                        relationId,
                        placement.ColumnIndex,
                        row.TargetRowIndex,
                        rowOffsetEmu,
                        payload,
                        nextObjectId()));

                    rowOffsetEmu += PixelsToEmu(payload.HeightPx + 4);
                }
            }
        }

        worksheetDrawing.Save();
    }

    private static DrawingsPart EnsureDrawingsPart(WorksheetPart worksheetPart)
    {
        if (worksheetPart.DrawingsPart != null)
        {
            return worksheetPart.DrawingsPart;
        }

        var drawingsPart = worksheetPart.AddNewPart<DrawingsPart>();
        drawingsPart.WorksheetDrawing = new Xdr.WorksheetDrawing();

        var relationId = worksheetPart.GetIdOfPart(drawingsPart);
        var drawing = worksheetPart.Worksheet.Elements<S.Drawing>().FirstOrDefault();
        if (drawing == null)
        {
            drawing = new S.Drawing { Id = relationId };
            worksheetPart.Worksheet.Append(drawing);
        }
        else
        {
            drawing.Id = relationId;
        }

        return drawingsPart;
    }

    private static Xdr.OneCellAnchor CreateAnchor(
        string relationId,
        int columnIndex,
        uint rowIndex,
        long rowOffsetEmu,
        ImagePayload payload,
        uint objectId)
    {
        var picture = new Xdr.Picture(
            new Xdr.NonVisualPictureProperties(
                new Xdr.NonVisualDrawingProperties
                {
                    Id = objectId,
                    Name = "Image " + objectId.ToString(CultureInfo.InvariantCulture)
                },
                new Xdr.NonVisualPictureDrawingProperties(
                    new A.PictureLocks { NoChangeAspect = true })),
            new Xdr.BlipFill(
                new A.Blip { Embed = relationId, CompressionState = A.BlipCompressionValues.Print },
                new A.Stretch(new A.FillRectangle())),
            new Xdr.ShapeProperties(
                new A.Transform2D(
                    new A.Offset { X = 0L, Y = 0L },
                    new A.Extents { Cx = PixelsToEmu(payload.WidthPx), Cy = PixelsToEmu(payload.HeightPx) }),
                new A.PresetGeometry(new A.AdjustValueList())
                {
                    Preset = A.ShapeTypeValues.Rectangle
                }));

        return new Xdr.OneCellAnchor(
            new Xdr.FromMarker(
                new Xdr.ColumnId((columnIndex - 1).ToString(CultureInfo.InvariantCulture)),
                new Xdr.ColumnOffset("0"),
                new Xdr.RowId((rowIndex - 1U).ToString(CultureInfo.InvariantCulture)),
                new Xdr.RowOffset(rowOffsetEmu.ToString(CultureInfo.InvariantCulture))),
            new Xdr.Extent
            {
                Cx = PixelsToEmu(payload.WidthPx),
                Cy = PixelsToEmu(payload.HeightPx)
            },
            picture,
            new Xdr.ClientData());
    }

    private static long PixelsToEmu(int pixels)
    {
        var safePixels = pixels <= 0 ? 1 : pixels;
        return safePixels * EmusPerPixel;
    }
}

internal sealed class SpreadsheetRowMapping
{
    private readonly Dictionary<uint, List<uint>> _rowsBySource = new Dictionary<uint, List<uint>>();
    private readonly Dictionary<string, Dictionary<uint, uint>> _rowsByScope = new Dictionary<string, Dictionary<uint, uint>>(StringComparer.Ordinal);
    private readonly string _rootScopeId;

    public SpreadsheetRowMapping(IEnumerable<RenderedSpreadsheetRow> rows, string rootScopeId)
    {
        _rootScopeId = rootScopeId;

        foreach (var row in rows)
        {
            if (!_rowsBySource.TryGetValue(row.SourceRowIndex, out var targets))
            {
                targets = new List<uint>();
                _rowsBySource[row.SourceRowIndex] = targets;
            }

            targets.Add(row.TargetRowIndex);

            if (!_rowsByScope.TryGetValue(row.ScopeId, out var scopeRows))
            {
                scopeRows = new Dictionary<uint, uint>();
                _rowsByScope[row.ScopeId] = scopeRows;
            }

            if (!scopeRows.ContainsKey(row.SourceRowIndex))
            {
                scopeRows[row.SourceRowIndex] = row.TargetRowIndex;
            }
        }
    }

    public IEnumerable<string> ScopeIds => _rowsByScope.Keys;

    public bool TryResolveRow(string scopeId, uint sourceRowIndex, out uint targetRowIndex)
    {
        if (_rowsByScope.TryGetValue(scopeId, out var scopeRows) && scopeRows.TryGetValue(sourceRowIndex, out targetRowIndex))
        {
            return true;
        }

        if (TryGetUniqueGlobalRow(sourceRowIndex, out targetRowIndex))
        {
            return true;
        }

        if (_rowsBySource.TryGetValue(sourceRowIndex, out var rows) && rows.Count > 0)
        {
            targetRowIndex = rows[0];
            return true;
        }

        targetRowIndex = 0U;
        return false;
    }

    public bool TryResolveRange(string scopeId, uint startSourceRow, uint endSourceRow, out uint startTargetRow, out uint endTargetRow)
    {
        if (TryResolveRangeInScope(scopeId, startSourceRow, endSourceRow, out startTargetRow, out endTargetRow))
        {
            return true;
        }

        return TryResolveRangeGlobal(startSourceRow, endSourceRow, out startTargetRow, out endTargetRow);
    }

    public bool TryResolveRangeInScope(string scopeId, uint startSourceRow, uint endSourceRow, out uint startTargetRow, out uint endTargetRow)
    {
        if (!_rowsByScope.TryGetValue(scopeId, out var scopeRows))
        {
            startTargetRow = 0U;
            endTargetRow = 0U;
            return false;
        }

        var ascendingStart = Math.Min(startSourceRow, endSourceRow);
        var ascendingEnd = Math.Max(startSourceRow, endSourceRow);
        var collectedRows = new List<uint>();
        var usedScopeRow = false;

        for (uint sourceRow = ascendingStart; sourceRow <= ascendingEnd; sourceRow++)
        {
            if (scopeRows.TryGetValue(sourceRow, out var scopedRow))
            {
                collectedRows.Add(scopedRow);
                usedScopeRow = true;
                continue;
            }

            if (TryGetUniqueGlobalRow(sourceRow, out var uniqueRow))
            {
                collectedRows.Add(uniqueRow);
                continue;
            }

            startTargetRow = 0U;
            endTargetRow = 0U;
            return false;
        }

        if (!usedScopeRow && !string.Equals(scopeId, _rootScopeId, StringComparison.Ordinal))
        {
            startTargetRow = 0U;
            endTargetRow = 0U;
            return false;
        }

        ApplyRangeOrder(startSourceRow, endSourceRow, collectedRows, out startTargetRow, out endTargetRow);
        return true;
    }

    public bool TryResolveRangeGlobal(uint startSourceRow, uint endSourceRow, out uint startTargetRow, out uint endTargetRow)
    {
        var ascendingStart = Math.Min(startSourceRow, endSourceRow);
        var ascendingEnd = Math.Max(startSourceRow, endSourceRow);
        var collectedRows = new List<uint>();

        for (uint sourceRow = ascendingStart; sourceRow <= ascendingEnd; sourceRow++)
        {
            if (_rowsBySource.TryGetValue(sourceRow, out var mappedRows))
            {
                collectedRows.AddRange(mappedRows);
            }
        }

        if (collectedRows.Count == 0)
        {
            startTargetRow = 0U;
            endTargetRow = 0U;
            return false;
        }

        ApplyRangeOrder(startSourceRow, endSourceRow, collectedRows, out startTargetRow, out endTargetRow);
        return true;
    }

    private bool TryGetUniqueGlobalRow(uint sourceRowIndex, out uint targetRowIndex)
    {
        if (_rowsBySource.TryGetValue(sourceRowIndex, out var rows) && rows.Count == 1)
        {
            targetRowIndex = rows[0];
            return true;
        }

        targetRowIndex = 0U;
        return false;
    }

    private static void ApplyRangeOrder(uint startSourceRow, uint endSourceRow, IReadOnlyCollection<uint> rows, out uint startTargetRow, out uint endTargetRow)
    {
        var minRow = rows.Min();
        var maxRow = rows.Max();
        if (startSourceRow <= endSourceRow)
        {
            startTargetRow = minRow;
            endTargetRow = maxRow;
            return;
        }

        startTargetRow = maxRow;
        endTargetRow = minRow;
    }
}

internal sealed class RenderedSpreadsheetRow
{
    public RenderedSpreadsheetRow(S.Row row, uint sourceRowIndex, string scopeId)
    {
        Row = row;
        SourceRowIndex = sourceRowIndex;
        ScopeId = scopeId;
        MediaPlacements = new List<SpreadsheetMediaPlacement>();
    }

    public S.Row Row { get; }

    public uint SourceRowIndex { get; }

    public string ScopeId { get; }

    public uint TargetRowIndex { get; set; }

    public List<SpreadsheetMediaPlacement> MediaPlacements { get; }
}

internal readonly struct SpreadsheetSourceRow
{
    public SpreadsheetSourceRow(S.Row row, uint sourceRowIndex)
    {
        Row = row;
        SourceRowIndex = sourceRowIndex;
    }

    public S.Row Row { get; }

    public uint SourceRowIndex { get; }
}

internal readonly struct SpreadsheetMediaPlacement
{
    public SpreadsheetMediaPlacement(int columnIndex, IReadOnlyList<ImagePayload> payloads)
    {
        ColumnIndex = columnIndex;
        Payloads = payloads;
    }

    public int ColumnIndex { get; }

    public IReadOnlyList<ImagePayload> Payloads { get; }
}

internal readonly struct SpreadsheetCellReference
{
    private SpreadsheetCellReference(string sheetPrefix, string columnName, bool columnAbsolute, uint rowIndex, bool rowAbsolute)
    {
        SheetPrefix = sheetPrefix;
        ColumnName = columnName;
        ColumnAbsolute = columnAbsolute;
        RowIndex = rowIndex;
        RowAbsolute = rowAbsolute;
    }

    public string SheetPrefix { get; }

    public string ColumnName { get; }

    public bool ColumnAbsolute { get; }

    public uint RowIndex { get; }

    public bool RowAbsolute { get; }

    public string WithRow(uint rowIndex)
    {
        return SheetPrefix
            + (ColumnAbsolute ? "$" : string.Empty)
            + ColumnName
            + (RowAbsolute ? "$" : string.Empty)
            + rowIndex.ToString(CultureInfo.InvariantCulture);
    }

    public static bool TryParse(string reference, out SpreadsheetCellReference cellReference)
    {
        cellReference = default;
        if (string.IsNullOrWhiteSpace(reference))
        {
            return false;
        }

        var trimmed = reference.Trim();
        var bangIndex = trimmed.LastIndexOf('!');
        var sheetPrefix = bangIndex >= 0 ? trimmed.Substring(0, bangIndex + 1) : string.Empty;
        var cellPart = bangIndex >= 0 ? trimmed.Substring(bangIndex + 1) : trimmed;
        if (cellPart.Length < 2)
        {
            return false;
        }

        var index = 0;
        var columnAbsolute = false;
        if (cellPart[index] == '$')
        {
            columnAbsolute = true;
            index++;
        }

        var columnStart = index;
        while (index < cellPart.Length && char.IsLetter(cellPart[index]))
        {
            index++;
        }

        if (index == columnStart)
        {
            return false;
        }

        var columnName = cellPart.Substring(columnStart, index - columnStart).ToUpperInvariant();
        var rowAbsolute = false;
        if (index < cellPart.Length && cellPart[index] == '$')
        {
            rowAbsolute = true;
            index++;
        }

        if (index >= cellPart.Length)
        {
            return false;
        }

        if (!uint.TryParse(cellPart.Substring(index), NumberStyles.Integer, CultureInfo.InvariantCulture, out var rowIndex))
        {
            return false;
        }

        cellReference = new SpreadsheetCellReference(sheetPrefix, columnName, columnAbsolute, rowIndex, rowAbsolute);
        return true;
    }
}

internal readonly struct SpreadsheetRangeReference
{
    private SpreadsheetRangeReference(SpreadsheetCellReference start, SpreadsheetCellReference end)
    {
        Start = start;
        End = end;
    }

    public SpreadsheetCellReference Start { get; }

    public SpreadsheetCellReference End { get; }

    public string WithRows(uint startRowIndex, uint endRowIndex)
    {
        return Start.WithRow(startRowIndex) + ":" + End.WithRow(endRowIndex);
    }

    public static bool TryParse(string reference, out SpreadsheetRangeReference range)
    {
        range = default;
        if (string.IsNullOrWhiteSpace(reference))
        {
            return false;
        }

        var separatorIndex = reference.IndexOf(':');
        if (separatorIndex <= 0 || separatorIndex >= reference.Length - 1)
        {
            return false;
        }

        if (!SpreadsheetCellReference.TryParse(reference.Substring(0, separatorIndex), out var start)
            || !SpreadsheetCellReference.TryParse(reference.Substring(separatorIndex + 1), out var end))
        {
            return false;
        }

        range = new SpreadsheetRangeReference(start, end);
        return true;
    }
}
