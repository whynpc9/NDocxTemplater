using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using JArray = System.Text.Json.Nodes.JsonArray;
using JObject = System.Text.Json.Nodes.JsonObject;
using JToken = System.Text.Json.Nodes.JsonNode;

namespace NDocxTemplater;

internal sealed class TemplateContext
{
    public TemplateContext(JToken? current, JToken root, TemplateContext? parent, RenderOptions? options = null)
    {
        Current = current;
        Root = root;
        Parent = parent;
        Options = options ?? parent?.Options ?? RenderOptions.Default;
    }

    public JToken? Current { get; }

    public JToken Root { get; }

    public TemplateContext? Parent { get; }

    public RenderOptions Options { get; }

    public void ReportWarning(string code, string message, string? expression)
    {
        Options.WarningHandler?.Invoke(new RenderWarning(code, message, expression));
    }
}

internal static class ExpressionEvaluator
{
    public static JToken? Evaluate(string expression, TemplateContext context)
    {
        return Evaluate(expression, context, out _);
    }

    public static JToken? Evaluate(string expression, TemplateContext context, out bool resolved)
    {
        var steps = expression.Split(new[] { '|' }, StringSplitOptions.RemoveEmptyEntries)
            .Select(static part => part.Trim())
            .Where(static part => part.Length > 0)
            .ToList();

        if (steps.Count == 0)
        {
            resolved = false;
            return null;
        }

        var value = PathResolver.Resolve(steps[0], context, out resolved);
        if (!resolved)
        {
            HandleMissingValue(expression, context);
        }

        for (var index = 1; index < steps.Count; index++)
        {
            value = ApplyOperation(value, steps[index], context);
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
        return ToText(token, CultureInfo.InvariantCulture);
    }

    public static string ToText(JToken? token, TemplateContext context)
    {
        return ToText(token, context.Options.Culture);
    }

    private static string ToText(JToken? token, CultureInfo culture)
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
            return decimalValue.ToString(null, culture);
        }

        if (JsonNodeHelpers.TryGetDateTime(token, out var dateValue))
        {
            return dateValue.ToString("O", culture);
        }

        return token!.ToJsonString();
    }

    private static void HandleMissingValue(string expression, TemplateContext context)
    {
        var message = string.Format(
            CultureInfo.InvariantCulture,
            "Expression '{0}' could not be resolved from the current template context.",
            expression);

        if (context.Options.MissingValueBehavior == MissingValueBehavior.Throw)
        {
            throw new InvalidOperationException(message);
        }

        context.ReportWarning("MissingValue", message, expression);
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
            case "first":
                return ApplyFirst(value);
            case "last":
                return ApplyLast(value);
            case "nth":
                return ApplyNth(value, parts);
            case "at":
                return ApplyAt(value, parts);
            case "get":
            case "pick":
                return ApplyGet(value, parts);
            case "maxby":
                return ApplyExtremaBy(value, parts, pickMax: true);
            case "minby":
                return ApplyExtremaBy(value, parts, pickMax: false);
            case "count":
                return JsonValue.Create(Count(value));
            case "if":
                return ApplyInlineIf(value, parts);
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

    private static JToken? ApplyFirst(JToken? value)
    {
        if (!(value is JArray array))
        {
            return value;
        }

        if (array.Count == 0)
        {
            return null;
        }

        return JsonNodeHelpers.DeepClone(array[0]);
    }

    private static JToken? ApplyLast(JToken? value)
    {
        if (!(value is JArray array))
        {
            return value;
        }

        if (array.Count == 0)
        {
            return null;
        }

        return JsonNodeHelpers.DeepClone(array[array.Count - 1]);
    }

    private static JToken? ApplyNth(JToken? value, IReadOnlyList<string> parts)
    {
        if (!(value is JArray))
        {
            return value;
        }

        if (parts.Count < 2 || !int.TryParse(parts[1].Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out var rank))
        {
            throw new InvalidOperationException("nth operation requires integer rank (1-based): nth:N.");
        }

        if (rank <= 0)
        {
            throw new InvalidOperationException("nth operation rank must be greater than zero: nth:N.");
        }

        return ApplyArrayIndex(value, rank - 1);
    }

    private static JToken? ApplyAt(JToken? value, IReadOnlyList<string> parts)
    {
        if (!(value is JArray))
        {
            return value;
        }

        if (parts.Count < 2 || !int.TryParse(parts[1].Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out var index))
        {
            throw new InvalidOperationException("at operation requires integer index (0-based): at:index.");
        }

        return ApplyArrayIndex(value, index);
    }

    private static JToken? ApplyArrayIndex(JToken? value, int index)
    {
        if (!(value is JArray array))
        {
            return value;
        }

        var normalizedIndex = index < 0 ? array.Count + index : index;
        if (normalizedIndex < 0 || normalizedIndex >= array.Count)
        {
            return null;
        }

        return JsonNodeHelpers.DeepClone(array[normalizedIndex]);
    }

    private static JToken? ApplyGet(JToken? value, IReadOnlyList<string> parts)
    {
        if (parts.Count < 2)
        {
            throw new InvalidOperationException("get operation requires a path: get:path.");
        }

        var path = string.Join(":", parts.Skip(1)).Trim();
        if (path.Length == 0 || path == ".")
        {
            return value == null ? null : JsonNodeHelpers.DeepClone(value);
        }

        return JsonNodeHelpers.DeepClone(PathResolver.ResolveFrom(value, path));
    }

    private static JToken? ApplyExtremaBy(JToken? value, IReadOnlyList<string> parts, bool pickMax)
    {
        if (!(value is JArray array))
        {
            return value;
        }

        if (parts.Count < 2)
        {
            throw new InvalidOperationException(
                pickMax
                    ? "maxby operation requires key path: maxby:key."
                    : "minby operation requires key path: minby:key.");
        }

        if (array.Count == 0)
        {
            return null;
        }

        var keyPath = string.Join(":", parts.Skip(1)).Trim();
        if (keyPath.Length == 0)
        {
            throw new InvalidOperationException(
                pickMax
                    ? "maxby operation requires key path: maxby:key."
                    : "minby operation requires key path: minby:key.");
        }

        JToken? bestItem = null;
        JToken? bestKey = null;

        foreach (var item in array)
        {
            var currentKey = PathResolver.ResolveFrom(item, keyPath);
            if (bestItem == null)
            {
                bestItem = item;
                bestKey = currentKey;
                continue;
            }

            var comparison = CompareTokens(currentKey, bestKey);
            if ((pickMax && comparison > 0) || (!pickMax && comparison < 0))
            {
                bestItem = item;
                bestKey = currentKey;
            }
        }

        return JsonNodeHelpers.DeepClone(bestItem);
    }

    private static JToken ApplyInlineIf(JToken? value, IReadOnlyList<string> parts)
    {
        if (parts.Count < 2)
        {
            throw new InvalidOperationException("if operation requires at least true branch text: if:trueText[:falseText].");
        }

        var whenTrue = parts[1];
        var whenFalse = parts.Count >= 3
            ? string.Join(":", parts.Skip(2))
            : string.Empty;

        return JsonValue.Create(IsTruthy(value) ? whenTrue : whenFalse)!;
    }

    private static JToken ApplyFormat(JToken? value, IReadOnlyList<string> parts, TemplateContext context)
    {
        if (parts.Count < 2)
        {
            throw new InvalidOperationException("format operation requires format kind: format:number:0.00.");
        }

        var formatKind = parts[1].Trim().ToLowerInvariant();
        var pattern = string.Join(":", parts.Skip(2)).Trim();

        switch (formatKind)
        {
            case "number":
            case "numeric":
                if (string.IsNullOrEmpty(pattern))
                {
                    throw new InvalidOperationException("format:number requires pattern: format:number:0.00.");
                }

                if (TryGetDecimal(value, out var decimalValue))
                {
                    return JsonValue.Create(decimalValue.ToString(pattern, context.Options.Culture))!;
                }
                break;
            case "percent":
            case "percentage":
                if (TryGetDecimal(value, out var percentValue))
                {
                    var percentPattern = string.IsNullOrEmpty(pattern) ? "0.##%" : EnsureSuffixPattern(pattern, "%");
                    return JsonValue.Create(percentValue.ToString(percentPattern, context.Options.Culture))!;
                }
                break;
            case "permille":
            case "per-mille":
            case "per_mille":
                if (TryGetDecimal(value, out var permilleValue))
                {
                    var permillePattern = string.IsNullOrEmpty(pattern) ? "0.##‰" : EnsureSuffixPattern(pattern, "‰");
                    return JsonValue.Create(permilleValue.ToString(permillePattern, context.Options.Culture))!;
                }
                break;
            case "date":
            case "datetime":
            case "time":
                if (string.IsNullOrEmpty(pattern))
                {
                    throw new InvalidOperationException("format:date requires pattern: format:date:yyyy-MM-dd.");
                }

                if (TryGetDateTime(value, out var dateValue))
                {
                    return JsonValue.Create(dateValue.ToString(pattern, context.Options.Culture))!;
                }
                break;
            default:
                throw new InvalidOperationException(
                    string.Format(
                        CultureInfo.InvariantCulture,
                        "Unsupported format kind '{0}'.",
                        formatKind));
        }

        return JsonValue.Create(ToText(value, context))!;
    }

    private static string EnsureSuffixPattern(string pattern, string suffix)
    {
        var trimmedPattern = pattern.Trim();
        if (trimmedPattern.Length == 0)
        {
            return suffix;
        }

        return trimmedPattern.IndexOf(suffix, StringComparison.Ordinal) >= 0
            ? trimmedPattern
            : trimmedPattern + suffix;
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
        return Resolve(pathExpression, context, out _);
    }

    public static JToken? Resolve(string pathExpression, TemplateContext context, out bool resolved)
    {
        if (string.IsNullOrWhiteSpace(pathExpression))
        {
            resolved = false;
            return null;
        }

        var path = pathExpression.Trim();
        if (path == ".")
        {
            resolved = true;
            return context.Current;
        }

        if (path == "$")
        {
            resolved = true;
            return context.Root;
        }

        if (path.StartsWith("$.", StringComparison.Ordinal))
        {
            return ResolveFrom(context.Root, path.Substring(2), out resolved);
        }

        var currentResult = ResolveFrom(context.Current, path, out var currentResolved);
        if (currentResolved)
        {
            resolved = true;
            return currentResult;
        }

        var parent = context.Parent;
        while (parent != null)
        {
            var parentResult = ResolveFrom(parent.Current, path, out var parentResolved);
            if (parentResolved)
            {
                resolved = true;
                return parentResult;
            }

            parent = parent.Parent;
        }

        return ResolveFrom(context.Root, path, out resolved);
    }

    public static JToken? ResolveFrom(JToken? start, string path)
    {
        return ResolveFrom(start, path, out _);
    }

    public static JToken? ResolveFrom(JToken? start, string path, out bool resolved)
    {
        if (start == null)
        {
            resolved = false;
            return null;
        }

        if (string.IsNullOrWhiteSpace(path))
        {
            resolved = true;
            return start;
        }

        var cursor = start;
        foreach (var segment in ParsePath(path.Trim()))
        {
            if (segment.IsIndex)
            {
                if (!(cursor is JArray array) || segment.Index < 0 || segment.Index >= array.Count)
                {
                    resolved = false;
                    return null;
                }

                cursor = array[segment.Index];
                continue;
            }

            if (!(cursor is JObject obj) || !JsonNodeHelpers.TryGetPropertyValue(obj, segment.Name, out var propertyValue))
            {
                resolved = false;
                return null;
            }

            cursor = propertyValue;
        }

        resolved = true;
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
