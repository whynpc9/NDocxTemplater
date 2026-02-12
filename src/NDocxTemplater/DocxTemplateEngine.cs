using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

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

        var rootData = JToken.Parse(jsonData);

        using (var document = WordprocessingDocument.Open(outputStream, true))
        {
            if (document.MainDocumentPart?.Document?.Body == null)
            {
                throw new InvalidOperationException("The DOCX template does not contain a valid document body.");
            }

            var renderer = new OpenXmlTemplateRenderer(rootData);
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

    public OpenXmlTemplateRenderer(JToken rootData)
    {
        _rootData = rootData;
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

    private static void RenderBlock(
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

    private static void RenderElement(OpenXmlElement element, TemplateContext context)
    {
        if (element is OpenXmlCompositeElement composite)
        {
            var renderer = new OpenXmlTemplateRenderer(context.Root);
            renderer.RenderContainer(composite, context);
        }

        ReplaceInlineTags(element, context);
    }

    private static void ReplaceInlineTags(OpenXmlElement element, TemplateContext context)
    {
        foreach (var textNode in element.Descendants<Text>())
        {
            if (string.IsNullOrEmpty(textNode.Text))
            {
                continue;
            }

            textNode.Text = TagPatterns.InlineTagRegex.Replace(textNode.Text, match =>
            {
                var expression = match.Groups[1].Value.Trim();

                if (ControlMarker.IsControlToken(expression))
                {
                    return string.Empty;
                }

                var resolved = ExpressionEvaluator.Evaluate(expression, context);
                return ExpressionEvaluator.ToText(resolved);
            });
        }
    }
}

internal sealed class TemplateContext
{
    public TemplateContext(JToken current, JToken root, TemplateContext? parent)
    {
        Current = current;
        Root = root;
        Parent = parent;
    }

    public JToken Current { get; }

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
            value = ApplyOperation(value, steps[index], context);
        }

        return value;
    }

    public static IEnumerable<JToken> ToLoopItems(JToken? token)
    {
        if (token == null || token.Type == JTokenType.Null)
        {
            return Enumerable.Empty<JToken>();
        }

        if (token is JArray array)
        {
            return array.Children<JToken>().ToList();
        }

        if (IsTruthy(token))
        {
            return new[] { token };
        }

        return Enumerable.Empty<JToken>();
    }

    public static bool IsTruthy(JToken? token)
    {
        if (token == null || token.Type == JTokenType.Null || token.Type == JTokenType.Undefined)
        {
            return false;
        }

        switch (token.Type)
        {
            case JTokenType.Boolean:
                return token.Value<bool>();
            case JTokenType.String:
                return !string.IsNullOrWhiteSpace(token.Value<string>());
            case JTokenType.Integer:
            case JTokenType.Float:
                return Math.Abs(token.Value<double>()) > double.Epsilon;
            case JTokenType.Array:
                return token.Any();
            case JTokenType.Object:
                return token.HasValues;
            default:
                return true;
        }
    }

    public static string ToText(JToken? token)
    {
        if (token == null || token.Type == JTokenType.Null || token.Type == JTokenType.Undefined)
        {
            return string.Empty;
        }

        if (token is JValue value)
        {
            if (value.Value is IFormattable formattable)
            {
                return formattable.ToString(null, CultureInfo.InvariantCulture);
            }

            return value.Value?.ToString() ?? string.Empty;
        }

        return token.ToString(Formatting.None);
    }

    private static JToken? ApplyOperation(JToken? value, string operation, TemplateContext context)
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
                return new JValue(Count(value));
            case "format":
                return ApplyFormat(value, parts, context);
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

        var sorted = sourceArray.Children<JToken>().ToList();
        sorted.Sort((left, right) => CompareTokens(
            PathResolver.ResolveFrom(left, keyPath),
            PathResolver.ResolveFrom(right, keyPath)));

        if (descending)
        {
            sorted.Reverse();
        }

        return new JArray(sorted);
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

        if (takeCount <= 0)
        {
            return new JArray();
        }

        return new JArray(sourceArray.Children<JToken>().Take(takeCount));
    }

    private static JToken ApplyFormat(JToken? value, IReadOnlyList<string> parts, TemplateContext context)
    {
        if (parts.Count < 3)
        {
            throw new InvalidOperationException("format operation requires kind and pattern: format:number:0.00.");
        }

        var formatKind = parts[1].Trim().ToLowerInvariant();
        var pattern = string.Join(":", parts.Skip(2)).Trim();

        if (string.IsNullOrEmpty(pattern))
        {
            return new JValue(string.Empty);
        }

        switch (formatKind)
        {
            case "number":
            case "numeric":
                if (TryGetDecimal(value, out var decimalValue))
                {
                    return new JValue(decimalValue.ToString(pattern, CultureInfo.InvariantCulture));
                }
                break;
            case "date":
            case "datetime":
            case "time":
                if (TryGetDateTime(value, out var dateValue))
                {
                    return new JValue(dateValue.ToString(pattern, CultureInfo.InvariantCulture));
                }
                break;
            default:
                throw new InvalidOperationException(
                    string.Format(
                        CultureInfo.InvariantCulture,
                        "Unsupported format kind '{0}'.",
                        formatKind));
        }

        return new JValue(ToText(value));
    }

    private static int Count(JToken? value)
    {
        if (value == null || value.Type == JTokenType.Null || value.Type == JTokenType.Undefined)
        {
            return 0;
        }

        switch (value.Type)
        {
            case JTokenType.Array:
                return value.Count();
            case JTokenType.Object:
                return value.Children<JProperty>().Count();
            case JTokenType.String:
                return value.Value<string>()?.Length ?? 0;
            default:
                return 1;
        }
    }

    private static int CompareTokens(JToken? left, JToken? right)
    {
        if (left == null || left.Type == JTokenType.Null || left.Type == JTokenType.Undefined)
        {
            return right == null || right.Type == JTokenType.Null || right.Type == JTokenType.Undefined ? 0 : -1;
        }

        if (right == null || right.Type == JTokenType.Null || right.Type == JTokenType.Undefined)
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
        if (token == null || token.Type == JTokenType.Null || token.Type == JTokenType.Undefined)
        {
            return false;
        }

        if (token is JValue jValue)
        {
            switch (jValue.Type)
            {
                case JTokenType.Integer:
                case JTokenType.Float:
                    value = Convert.ToDecimal(jValue.Value, CultureInfo.InvariantCulture);
                    return true;
            }
        }

        return decimal.TryParse(ToText(token), NumberStyles.Any, CultureInfo.InvariantCulture, out value);
    }

    private static bool TryGetDateTime(JToken? token, out DateTime value)
    {
        value = default;
        if (token == null || token.Type == JTokenType.Null || token.Type == JTokenType.Undefined)
        {
            return false;
        }

        if (token is JValue jValue && jValue.Value is DateTime date)
        {
            value = date;
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

            if (!(cursor is JObject obj) || !obj.TryGetValue(segment.Name, StringComparison.Ordinal, out var propertyValue))
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
