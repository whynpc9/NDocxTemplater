using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Text.Json;
using System.Text.Json.Nodes;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using JArray = System.Text.Json.Nodes.JsonArray;
using JObject = System.Text.Json.Nodes.JsonObject;
using JToken = System.Text.Json.Nodes.JsonNode;
using JValue = System.Text.Json.Nodes.JsonValue;

namespace NDocxTemplater;

public sealed class DocxTemplateEngine
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

        using (var document = WordprocessingDocument.Open(outputStream, true))
        {
            if (document.MainDocumentPart?.Document?.Body == null)
            {
                throw new InvalidOperationException("The DOCX template does not contain a valid document body.");
            }

            var renderer = new OpenXmlTemplateRenderer(rootData, document.MainDocumentPart);
            var rootContext = new TemplateContext(rootData, rootData, null);
            renderer.RenderContainer(document.MainDocumentPart.Document.Body, rootContext);
            document.MainDocumentPart.Document.Save();
        }

        outputStream.Position = 0;
    }
}

internal sealed class OpenXmlTemplateRenderer
{
    private readonly JToken _rootData;
    private readonly MainDocumentPart _mainDocumentPart;
    private uint _imageIdCounter = 1;

    public OpenXmlTemplateRenderer(JToken rootData, MainDocumentPart mainDocumentPart)
    {
        _rootData = rootData;
        _mainDocumentPart = mainDocumentPart;
    }

    public void RenderContainer(OpenXmlCompositeElement container, TemplateContext context)
    {
        var sourceChildren = container.ChildElements.Cast<OpenXmlElement>().ToList();
        var renderedChildren = new List<OpenXmlElement>();

        for (var index = 0; index < sourceChildren.Count; index++)
        {
            var candidate = sourceChildren[index];
            var marker = ControlMarker.TryParse(candidate);

            if (marker != null && marker.IsStart)
            {
                var endIndex = FindMatchingEnd(sourceChildren, index, marker);
                var blockTemplates = sourceChildren.Skip(index + 1).Take(endIndex - index - 1).ToList();

                if (marker.Kind == ControlMarkerKind.LoopStart)
                {
                    var loopData = ExpressionEvaluator.Evaluate(marker.Expression, context);
                    foreach (var item in ExpressionEvaluator.ToLoopItems(loopData))
                    {
                        var itemContext = new TemplateContext(item, _rootData, context);
                        RenderBlock(blockTemplates, renderedChildren, itemContext);
                    }
                }
                else if (marker.Kind == ControlMarkerKind.IfStart)
                {
                    var conditionValue = ExpressionEvaluator.Evaluate(marker.Expression, context);
                    if (ExpressionEvaluator.IsTruthy(conditionValue))
                    {
                        RenderBlock(blockTemplates, renderedChildren, context);
                    }
                }

                index = endIndex;
                continue;
            }

            if (marker != null && marker.IsEnd)
            {
                continue;
            }

            var cloned = candidate.CloneNode(true);
            RenderElement(cloned, context);
            renderedChildren.Add(cloned);
        }

        container.RemoveAllChildren();
        foreach (var rendered in renderedChildren)
        {
            container.AppendChild(rendered);
        }
    }

    private static int FindMatchingEnd(IReadOnlyList<OpenXmlElement> siblings, int startIndex, ControlMarker startMarker)
    {
        var depth = 0;

        for (var index = startIndex + 1; index < siblings.Count; index++)
        {
            var marker = ControlMarker.TryParse(siblings[index]);
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

    private void RenderBlock(
        IReadOnlyCollection<OpenXmlElement> blockTemplates,
        ICollection<OpenXmlElement> renderedChildren,
        TemplateContext context)
    {
        foreach (var blockTemplate in blockTemplates)
        {
            var clone = blockTemplate.CloneNode(true);
            RenderElement(clone, context);
            renderedChildren.Add(clone);
        }
    }

    private void RenderElement(OpenXmlElement element, TemplateContext context)
    {
        if (element is OpenXmlCompositeElement composite)
        {
            RenderContainer(composite, context);
        }

        if (element is Paragraph paragraph
            && ImageTemplateRenderer.TryRenderImageTag(paragraph, context, _mainDocumentPart, NextImageId))
        {
            return;
        }

        ReplaceInlineTags(element, context);
    }

    private uint NextImageId()
    {
        return _imageIdCounter++;
    }

    private static void ReplaceInlineTags(OpenXmlElement element, TemplateContext context)
    {
        if (element is Paragraph paragraph)
        {
            ReplaceInlineTagsInParagraph(paragraph, context);
            return;
        }

        foreach (var textNode in element.Descendants<Text>())
        {
            if (string.IsNullOrEmpty(textNode.Text))
            {
                continue;
            }

            textNode.Text = ReplaceInlineTagsInText(textNode.Text, context);
        }
    }

    private static void ReplaceInlineTagsInParagraph(Paragraph paragraph, TemplateContext context)
    {
        var textNodes = paragraph.Descendants<Text>().ToList();
        if (textNodes.Count == 0)
        {
            return;
        }

        if (textNodes.Count == 1)
        {
            var onlyText = textNodes[0];
            if (!string.IsNullOrEmpty(onlyText.Text))
            {
                onlyText.Text = ReplaceInlineTagsInText(onlyText.Text, context);
            }

            return;
        }

        var combinedText = string.Concat(textNodes.Select(static node => node.Text));
        if (string.IsNullOrEmpty(combinedText) || combinedText.IndexOf('{') < 0 || combinedText.IndexOf('}') < 0)
        {
            foreach (var textNode in textNodes)
            {
                if (!string.IsNullOrEmpty(textNode.Text))
                {
                    textNode.Text = ReplaceInlineTagsInText(textNode.Text, context);
                }
            }

            return;
        }

        var combinedReplaced = ReplaceInlineTagsInText(combinedText, context);
        var perNodeReplacedCombined = string.Concat(textNodes.Select(node => ReplaceInlineTagsInText(node.Text, context)));

        if (string.Equals(combinedReplaced, perNodeReplacedCombined, StringComparison.Ordinal))
        {
            for (var index = 0; index < textNodes.Count; index++)
            {
                textNodes[index].Text = ReplaceInlineTagsInText(textNodes[index].Text, context);
            }

            return;
        }

        textNodes[0].Text = combinedReplaced;
        for (var index = 1; index < textNodes.Count; index++)
        {
            textNodes[index].Text = string.Empty;
        }
    }

    private static string ReplaceInlineTagsInText(string text, TemplateContext context)
    {
        if (string.IsNullOrEmpty(text))
        {
            return text;
        }

        return TagPatterns.InlineTagRegex.Replace(text, match =>
        {
            var expression = match.Groups[1].Value.Trim();

            if (ControlMarker.IsControlToken(expression))
            {
                return string.Empty;
            }

            if (ImageTagParser.TryParseToken(expression, out _))
            {
                return match.Value;
            }

            var resolved = ExpressionEvaluator.Evaluate(expression, context);
            return ExpressionEvaluator.ToText(resolved);
        });
    }
}

internal sealed class TemplateContext
{
    public TemplateContext(JToken? current, JToken root, TemplateContext? parent)
    {
        Current = current;
        Root = root;
        Parent = parent;
    }

    public JToken? Current { get; }

    public JToken Root { get; }

    public TemplateContext? Parent { get; }
}

internal static class ExpressionEvaluator
{
    public static JToken? Evaluate(string expression, TemplateContext context)
    {
        var steps = expression.Split(new[] { '|' }, StringSplitOptions.RemoveEmptyEntries)
            .Select(static part => part.Trim())
            .Where(static part => part.Length > 0)
            .ToList();

        if (steps.Count == 0)
        {
            return null;
        }

        var value = PathResolver.Resolve(steps[0], context);

        for (var index = 1; index < steps.Count; index++)
        {
            value = ApplyOperation(value, steps[index]);
        }

        return value;
    }

    public static IEnumerable<JToken?> ToLoopItems(JToken? token)
    {
        if (JsonNodeHelpers.IsNull(token))
        {
            return Enumerable.Empty<JToken?>();
        }

        if (token is JArray array)
        {
            return array.Select(static item => item).ToList();
        }

        if (IsTruthy(token))
        {
            return new[] { token };
        }

        return Enumerable.Empty<JToken?>();
    }

    public static bool IsTruthy(JToken? token)
    {
        if (JsonNodeHelpers.IsNull(token))
        {
            return false;
        }

        if (JsonNodeHelpers.TryGetBoolean(token, out var boolValue))
        {
            return boolValue;
        }

        if (JsonNodeHelpers.TryGetString(token, out var stringValue))
        {
            return !string.IsNullOrWhiteSpace(stringValue);
        }

        if (JsonNodeHelpers.TryGetDouble(token, out var doubleValue))
        {
            return Math.Abs(doubleValue) > double.Epsilon;
        }

        if (token is JArray array)
        {
            return array.Count > 0;
        }

        if (token is JObject obj)
        {
            return obj.Count > 0;
        }

        return true;
    }

    public static string ToText(JToken? token)
    {
        if (JsonNodeHelpers.IsNull(token))
        {
            return string.Empty;
        }

        if (JsonNodeHelpers.TryGetString(token, out var stringValue))
        {
            return stringValue ?? string.Empty;
        }

        if (JsonNodeHelpers.TryGetBoolean(token, out var boolValue))
        {
            return boolValue ? "True" : "False";
        }

        if (JsonNodeHelpers.TryGetDecimal(token, out var decimalValue))
        {
            return decimalValue.ToString(null, CultureInfo.InvariantCulture);
        }

        if (JsonNodeHelpers.TryGetDateTime(token, out var dateValue))
        {
            return dateValue.ToString("O", CultureInfo.InvariantCulture);
        }

        return token!.ToJsonString();
    }

    private static JToken? ApplyOperation(JToken? value, string operation)
    {
        var parts = operation.Split(':');
        if (parts.Length == 0)
        {
            return value;
        }

        var command = parts[0].Trim().ToLowerInvariant();

        switch (command)
        {
            case "sort":
                return ApplySort(value, parts);
            case "take":
                return ApplyTake(value, parts);
            case "count":
                return JsonValue.Create(Count(value));
            case "format":
                return ApplyFormat(value, parts);
            default:
                throw new InvalidOperationException(
                    string.Format(
                        CultureInfo.InvariantCulture,
                        "Unsupported operation '{0}' in expression.",
                        operation));
        }
    }

    private static JToken? ApplySort(JToken? value, IReadOnlyList<string> parts)
    {
        if (!(value is JArray sourceArray))
        {
            return value;
        }

        if (parts.Count < 2)
        {
            throw new InvalidOperationException("sort operation requires key path: sort:key[:asc|desc].");
        }

        var keyPath = parts[1].Trim();
        var direction = parts.Count >= 3 ? parts[2].Trim() : "asc";
        var descending = string.Equals(direction, "desc", StringComparison.OrdinalIgnoreCase);

        var sorted = sourceArray.Select(static item => item).ToList();
        sorted.Sort((left, right) => CompareTokens(
            PathResolver.ResolveFrom(left, keyPath),
            PathResolver.ResolveFrom(right, keyPath)));

        if (descending)
        {
            sorted.Reverse();
        }

        var result = new JArray();
        foreach (var item in sorted)
        {
            result.Add(JsonNodeHelpers.DeepClone(item));
        }

        return result;
    }

    private static JToken? ApplyTake(JToken? value, IReadOnlyList<string> parts)
    {
        if (!(value is JArray sourceArray))
        {
            return value;
        }

        if (parts.Count < 2 || !int.TryParse(parts[1].Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out var takeCount))
        {
            throw new InvalidOperationException("take operation requires integer count: take:N.");
        }

        var result = new JArray();
        if (takeCount <= 0)
        {
            return result;
        }

        foreach (var item in sourceArray.Take(takeCount))
        {
            result.Add(JsonNodeHelpers.DeepClone(item));
        }

        return result;
    }

    private static JToken ApplyFormat(JToken? value, IReadOnlyList<string> parts)
    {
        if (parts.Count < 3)
        {
            throw new InvalidOperationException("format operation requires kind and pattern: format:number:0.00.");
        }

        var formatKind = parts[1].Trim().ToLowerInvariant();
        var pattern = string.Join(":", parts.Skip(2)).Trim();

        if (string.IsNullOrEmpty(pattern))
        {
            return JsonValue.Create(string.Empty)!;
        }

        switch (formatKind)
        {
            case "number":
            case "numeric":
                if (TryGetDecimal(value, out var decimalValue))
                {
                    return JsonValue.Create(decimalValue.ToString(pattern, CultureInfo.InvariantCulture))!;
                }
                break;
            case "date":
            case "datetime":
            case "time":
                if (TryGetDateTime(value, out var dateValue))
                {
                    return JsonValue.Create(dateValue.ToString(pattern, CultureInfo.InvariantCulture))!;
                }
                break;
            default:
                throw new InvalidOperationException(
                    string.Format(
                        CultureInfo.InvariantCulture,
                        "Unsupported format kind '{0}'.",
                        formatKind));
        }

        return JsonValue.Create(ToText(value))!;
    }

    private static int Count(JToken? value)
    {
        if (JsonNodeHelpers.IsNull(value))
        {
            return 0;
        }

        if (value is JArray array)
        {
            return array.Count;
        }

        if (value is JObject obj)
        {
            return obj.Count;
        }

        if (JsonNodeHelpers.TryGetString(value, out var stringValue))
        {
            return stringValue?.Length ?? 0;
        }

        return 1;
    }

    private static int CompareTokens(JToken? left, JToken? right)
    {
        if (JsonNodeHelpers.IsNull(left))
        {
            return JsonNodeHelpers.IsNull(right) ? 0 : -1;
        }

        if (JsonNodeHelpers.IsNull(right))
        {
            return 1;
        }

        if (TryGetDecimal(left, out var leftDecimal) && TryGetDecimal(right, out var rightDecimal))
        {
            return leftDecimal.CompareTo(rightDecimal);
        }

        if (TryGetDateTime(left, out var leftDate) && TryGetDateTime(right, out var rightDate))
        {
            return leftDate.CompareTo(rightDate);
        }

        var leftText = ToText(left);
        var rightText = ToText(right);
        return string.Compare(leftText, rightText, StringComparison.OrdinalIgnoreCase);
    }

    private static bool TryGetDecimal(JToken? token, out decimal value)
    {
        value = 0m;
        if (JsonNodeHelpers.IsNull(token))
        {
            return false;
        }

        if (JsonNodeHelpers.TryGetDecimal(token, out value))
        {
            return true;
        }

        return decimal.TryParse(ToText(token), NumberStyles.Any, CultureInfo.InvariantCulture, out value);
    }

    private static bool TryGetDateTime(JToken? token, out DateTime value)
    {
        value = default;
        if (JsonNodeHelpers.IsNull(token))
        {
            return false;
        }

        if (JsonNodeHelpers.TryGetDateTime(token, out value))
        {
            return true;
        }

        return DateTime.TryParse(ToText(token), CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind, out value)
            || DateTime.TryParse(ToText(token), CultureInfo.CurrentCulture, DateTimeStyles.None, out value);
    }
}

internal static class PathResolver
{
    public static JToken? Resolve(string pathExpression, TemplateContext context)
    {
        if (string.IsNullOrWhiteSpace(pathExpression))
        {
            return null;
        }

        var path = pathExpression.Trim();
        if (path == ".")
        {
            return context.Current;
        }

        if (path == "$")
        {
            return context.Root;
        }

        if (path.StartsWith("$.", StringComparison.Ordinal))
        {
            return ResolveFrom(context.Root, path.Substring(2));
        }

        var currentResult = ResolveFrom(context.Current, path);
        if (currentResult != null)
        {
            return currentResult;
        }

        var parent = context.Parent;
        while (parent != null)
        {
            var parentResult = ResolveFrom(parent.Current, path);
            if (parentResult != null)
            {
                return parentResult;
            }

            parent = parent.Parent;
        }

        return ResolveFrom(context.Root, path);
    }

    public static JToken? ResolveFrom(JToken? start, string path)
    {
        if (start == null || string.IsNullOrWhiteSpace(path))
        {
            return start;
        }

        var cursor = start;
        foreach (var segment in ParsePath(path.Trim()))
        {
            if (segment.IsIndex)
            {
                if (!(cursor is JArray array) || segment.Index < 0 || segment.Index >= array.Count)
                {
                    return null;
                }

                cursor = array[segment.Index];
                continue;
            }

            if (!(cursor is JObject obj) || !JsonNodeHelpers.TryGetPropertyValue(obj, segment.Name, out var propertyValue))
            {
                return null;
            }

            cursor = propertyValue;
        }

        return cursor;
    }

    private static IEnumerable<PathSegment> ParsePath(string path)
    {
        var segments = new List<PathSegment>();
        var index = 0;

        while (index < path.Length)
        {
            if (path[index] == '.')
            {
                index++;
                continue;
            }

            if (path[index] == '[')
            {
                var closingBracket = path.IndexOf(']', index + 1);
                if (closingBracket <= index + 1)
                {
                    throw new InvalidOperationException(
                        string.Format(CultureInfo.InvariantCulture, "Invalid path expression '{0}'.", path));
                }

                var indexText = path.Substring(index + 1, closingBracket - index - 1);
                if (!int.TryParse(indexText, NumberStyles.Integer, CultureInfo.InvariantCulture, out var itemIndex))
                {
                    throw new InvalidOperationException(
                        string.Format(CultureInfo.InvariantCulture, "Invalid array index '{0}' in path '{1}'.", indexText, path));
                }

                segments.Add(PathSegment.ForIndex(itemIndex));
                index = closingBracket + 1;
                continue;
            }

            var start = index;
            while (index < path.Length && path[index] != '.' && path[index] != '[')
            {
                index++;
            }

            var name = path.Substring(start, index - start).Trim();
            if (name.Length > 0)
            {
                segments.Add(PathSegment.ForName(name));
            }
        }

        return segments;
    }

    private struct PathSegment
    {
        private PathSegment(string name, int index, bool isIndex)
        {
            Name = name;
            Index = index;
            IsIndex = isIndex;
        }

        public string Name { get; }

        public int Index { get; }

        public bool IsIndex { get; }

        public static PathSegment ForName(string name)
        {
            return new PathSegment(name, -1, false);
        }

        public static PathSegment ForIndex(int index)
        {
            return new PathSegment(string.Empty, index, true);
        }
    }
}

internal static class JsonNodeHelpers
{
    public static bool IsNull(JToken? node)
    {
        return node == null;
    }

    public static bool TryGetPropertyValue(JObject obj, string name, out JToken? value)
    {
        foreach (var pair in obj)
        {
            if (string.Equals(pair.Key, name, StringComparison.Ordinal))
            {
                value = pair.Value;
                return true;
            }
        }

        value = null;
        return false;
    }

    public static bool TryGetString(JToken? node, out string? value)
    {
        if (node is JsonValue jsonValue)
        {
            try
            {
                value = jsonValue.GetValue<string>();
                return true;
            }
            catch
            {
            }
        }

        value = null;
        return false;
    }

    public static bool TryGetBoolean(JToken? node, out bool value)
    {
        if (node is JsonValue jsonValue)
        {
            try
            {
                value = jsonValue.GetValue<bool>();
                return true;
            }
            catch
            {
            }
        }

        value = default;
        return false;
    }

    public static bool TryGetDouble(JToken? node, out double value)
    {
        if (node is JsonValue jsonValue)
        {
            try
            {
                value = jsonValue.GetValue<double>();
                return true;
            }
            catch
            {
            }
        }

        value = default;
        return false;
    }

    public static bool TryGetDecimal(JToken? node, out decimal value)
    {
        if (node is JsonValue jsonValue)
        {
            try
            {
                value = jsonValue.GetValue<decimal>();
                return true;
            }
            catch
            {
            }

            try
            {
                value = Convert.ToDecimal(jsonValue.GetValue<double>(), CultureInfo.InvariantCulture);
                return true;
            }
            catch
            {
            }
        }

        value = default;
        return false;
    }

    public static bool TryGetDateTime(JToken? node, out DateTime value)
    {
        if (node is JsonValue jsonValue)
        {
            try
            {
                value = jsonValue.GetValue<DateTime>();
                return true;
            }
            catch
            {
            }
        }

        value = default;
        return false;
    }

    public static JToken? DeepClone(JToken? node)
    {
        if (node == null)
        {
            return null;
        }

        return JsonNode.Parse(node.ToJsonString());
    }
}

internal static class TagPatterns
{
    public static readonly Regex InlineTagRegex = new Regex(@"\{([^{}]+)\}", RegexOptions.Compiled);

    public static readonly Regex SingleTagRegex = new Regex(@"^\{([^{}]+)\}$", RegexOptions.Compiled);
}

internal enum ControlMarkerKind
{
    LoopStart,
    LoopEnd,
    IfStart,
    IfEnd
}

internal sealed class ControlMarker
{
    private ControlMarker(ControlMarkerKind kind, string expression, string rawToken)
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

    public static ControlMarker? TryParse(OpenXmlElement element)
    {
        var rawText = string.Concat(element.Descendants<Text>().Select(static text => text.Text)).Trim();
        if (rawText.Length == 0)
        {
            return null;
        }

        var fullTagMatch = TagPatterns.SingleTagRegex.Match(rawText);
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
            return expression.Length == 0 ? null : new ControlMarker(ControlMarkerKind.LoopStart, expression, "{" + token + "}");
        }

        if (token.StartsWith("/?", StringComparison.Ordinal))
        {
            var expression = token.Substring(2).Trim();
            return expression.Length == 0 ? null : new ControlMarker(ControlMarkerKind.IfEnd, expression, "{" + token + "}");
        }

        if (token.StartsWith("?", StringComparison.Ordinal))
        {
            var expression = token.Substring(1).Trim();
            return expression.Length == 0 ? null : new ControlMarker(ControlMarkerKind.IfStart, expression, "{" + token + "}");
        }

        if (token.StartsWith("/", StringComparison.Ordinal))
        {
            var expression = token.Substring(1).Trim();
            return expression.Length == 0 ? null : new ControlMarker(ControlMarkerKind.LoopEnd, expression, "{" + token + "}");
        }

        return null;
    }

    public static bool IsControlToken(string token)
    {
        if (string.IsNullOrWhiteSpace(token))
        {
            return false;
        }

        return token.StartsWith("#", StringComparison.Ordinal)
            || token.StartsWith("/?", StringComparison.Ordinal)
            || token.StartsWith("?", StringComparison.Ordinal)
            || token.StartsWith("/", StringComparison.Ordinal);
    }

    public static bool IsStartOfSameType(ControlMarkerKind blockStartKind, ControlMarkerKind candidateKind)
    {
        return (blockStartKind == ControlMarkerKind.LoopStart && candidateKind == ControlMarkerKind.LoopStart)
            || (blockStartKind == ControlMarkerKind.IfStart && candidateKind == ControlMarkerKind.IfStart);
    }

    public static bool IsEndOfSameType(ControlMarkerKind blockStartKind, ControlMarkerKind candidateKind)
    {
        return (blockStartKind == ControlMarkerKind.LoopStart && candidateKind == ControlMarkerKind.LoopEnd)
            || (blockStartKind == ControlMarkerKind.IfStart && candidateKind == ControlMarkerKind.IfEnd);
    }
}
