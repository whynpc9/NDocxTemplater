using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.Json.Nodes;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using JToken = System.Text.Json.Nodes.JsonNode;

namespace NDocxTemplater;

public sealed class DocxTemplateEngine
{
    public byte[] Render(byte[] templateBytes, string jsonData)
    {
        return Render(templateBytes, jsonData, null);
    }

    public byte[] Render(byte[] templateBytes, string jsonData, RenderOptions? options)
    {
        if (templateBytes == null)
        {
            throw new ArgumentNullException(nameof(templateBytes));
        }

        using (var templateStream = new MemoryStream(templateBytes, writable: false))
        using (var outputStream = new MemoryStream())
        {
            Render(templateStream, outputStream, jsonData, options);
            return outputStream.ToArray();
        }
    }

    public void Render(Stream templateStream, Stream outputStream, string jsonData)
    {
        Render(templateStream, outputStream, jsonData, null);
    }

    public void Render(Stream templateStream, Stream outputStream, string jsonData, RenderOptions? options)
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
            var rootContext = new TemplateContext(rootData, rootData, null, options);
            renderer.RenderDocument(document.MainDocumentPart.Document.Body, rootContext);
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

    public void RenderDocument(Body body, TemplateContext context)
    {
        RenderContainer(body, context);
        RenderHeaderFooterParts(context);
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
        if (element is TextBoxContent textBoxContent)
        {
            RenderContainer(textBoxContent, context);
            return;
        }

        if (element is Paragraph paragraph
            && ImageTemplateRenderer.TryRenderImageTag(paragraph, context, _mainDocumentPart, NextImageId))
        {
            return;
        }

        if (element is Paragraph paragraphElement)
        {
            RenderNestedTextBoxContents(paragraphElement, context);
            ReplaceInlineTags(paragraphElement, context);
            return;
        }

        if (element is OpenXmlCompositeElement composite)
        {
            RenderContainer(composite, context);
        }

        ReplaceInlineTags(element, context);
    }

    private void RenderHeaderFooterParts(TemplateContext context)
    {
        foreach (var headerPart in _mainDocumentPart.HeaderParts)
        {
            if (headerPart.Header == null)
            {
                continue;
            }

            RenderContainer(headerPart.Header, context);
            headerPart.Header.Save();
        }

        foreach (var footerPart in _mainDocumentPart.FooterParts)
        {
            if (footerPart.Footer == null)
            {
                continue;
            }

            RenderContainer(footerPart.Footer, context);
            footerPart.Footer.Save();
        }
    }

    private void RenderNestedTextBoxContents(OpenXmlElement element, TemplateContext context)
    {
        foreach (var textBoxContent in element.Descendants<TextBoxContent>().ToList())
        {
            RenderContainer(textBoxContent, context);
        }
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
            if (IsInsideTextBoxContent(textNode))
            {
                continue;
            }

            if (string.IsNullOrEmpty(textNode.Text))
            {
                continue;
            }

            textNode.Text = ReplaceInlineTagsInText(textNode.Text, context);
        }
    }

    private static void ReplaceInlineTagsInParagraph(Paragraph paragraph, TemplateContext context)
    {
        var textNodes = paragraph.Descendants<Text>()
            .Where(static text => !IsInsideTextBoxContent(text))
            .ToList();
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

    private static bool IsInsideTextBoxContent(OpenXmlElement element)
    {
        var parent = element.Parent;
        while (parent != null)
        {
            if (parent is TextBoxContent)
            {
                return true;
            }

            parent = parent.Parent;
        }

        return false;
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

            var resolved = ExpressionEvaluator.Evaluate(expression, context, out var expressionResolved);
            if (!expressionResolved && context.Options.MissingValueBehavior == MissingValueBehavior.KeepTag)
            {
                return match.Value;
            }

            return ExpressionEvaluator.ToText(resolved, context);
        });
    }
}
