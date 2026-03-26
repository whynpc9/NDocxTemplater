using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.Json.Nodes;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;
using JArray = System.Text.Json.Nodes.JsonArray;
using JObject = System.Text.Json.Nodes.JsonObject;
using JToken = System.Text.Json.Nodes.JsonNode;

namespace NDocxTemplater;

public sealed class PptxTemplateEngine
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

        using (var document = PresentationDocument.Open(outputStream, true))
        {
            if (document.PresentationPart?.Presentation == null)
            {
                throw new InvalidOperationException("The PPTX template does not contain a valid presentation.");
            }

            var renderer = new PptxPresentationRenderer(document.PresentationPart, rootData);
            renderer.Render();
            document.PresentationPart.Presentation.Save();
        }

        outputStream.Position = 0;
    }
}

internal sealed class PptxPresentationRenderer
{
    private readonly PresentationPart _presentationPart;
    private readonly JToken _rootData;

    public PptxPresentationRenderer(PresentationPart presentationPart, JToken rootData)
    {
        _presentationPart = presentationPart;
        _rootData = rootData;
    }

    public void Render()
    {
        var slideIdList = _presentationPart.Presentation.SlideIdList;
        if (slideIdList == null)
        {
            return;
        }

        var rootContext = new TemplateContext(_rootData, _rootData, null);
        var sourceSlides = slideIdList.Elements<P.SlideId>()
            .Select(slideId => new PptxSourceSlide(
                slideId,
                (SlidePart)_presentationPart.GetPartById(slideId.RelationshipId!)))
            .ToList();

        var renderedSlides = new List<SlidePart>();

        foreach (var sourceSlide in sourceSlides)
        {
            var contexts = ResolveSlideContexts(sourceSlide.SlidePart, rootContext);
            if (contexts.Count == 0)
            {
                continue;
            }

            var clonedSlides = contexts.Count > 1
                ? contexts.Skip(1)
                    .Select(_ => PptxSlideCloner.Clone(_presentationPart, sourceSlide.SlidePart))
                    .ToList()
                : new List<SlidePart>();

            for (var index = 0; index < contexts.Count; index++)
            {
                var targetSlidePart = index == 0
                    ? sourceSlide.SlidePart
                    : clonedSlides[index - 1];

                RenderSlide(targetSlidePart, contexts[index]);
                renderedSlides.Add(targetSlidePart);
            }
        }

        var retainedSlides = new HashSet<SlidePart>(renderedSlides);
        foreach (var sourceSlide in sourceSlides)
        {
            if (!retainedSlides.Contains(sourceSlide.SlidePart))
            {
                _presentationPart.DeletePart(sourceSlide.SlidePart);
            }
        }

        RebuildSlideIdList(slideIdList, renderedSlides);
    }

    private List<TemplateContext> ResolveSlideContexts(SlidePart slidePart, TemplateContext rootContext)
    {
        if (!PptxSlideTagParser.TryParse(slidePart.Slide, out var expression))
        {
            return new List<TemplateContext> { rootContext };
        }

        var value = ExpressionEvaluator.Evaluate(expression, rootContext);
        if (value is JArray array)
        {
            return array
                .Select(item => new TemplateContext(item, _rootData, rootContext))
                .ToList();
        }

        if (!ExpressionEvaluator.IsTruthy(value))
        {
            return new List<TemplateContext>();
        }

        if (value is JObject)
        {
            return new List<TemplateContext> { new TemplateContext(value, _rootData, rootContext) };
        }

        return new List<TemplateContext> { rootContext };
    }

    private void RenderSlide(SlidePart slidePart, TemplateContext context)
    {
        foreach (var paragraph in slidePart.Slide.Descendants<A.Paragraph>().ToList())
        {
            PptxParagraphRenderer.RenderParagraph(paragraph, context);
        }

        slidePart.Slide.Save();
    }

    private void RebuildSlideIdList(P.SlideIdList slideIdList, IReadOnlyList<SlidePart> slideParts)
    {
        slideIdList.RemoveAllChildren<P.SlideId>();

        uint nextSlideId = 256;
        foreach (var slidePart in slideParts)
        {
            slideIdList.Append(new P.SlideId
            {
                Id = nextSlideId++,
                RelationshipId = _presentationPart.GetIdOfPart(slidePart)
            });
        }
    }
}

internal static class PptxParagraphRenderer
{
    public static void RenderParagraph(A.Paragraph paragraph, TemplateContext context)
    {
        var rawText = GetParagraphText(paragraph);
        if (string.IsNullOrEmpty(rawText) || rawText.IndexOf('{') < 0 || rawText.IndexOf('}') < 0)
        {
            return;
        }

        var replaced = TagPatterns.InlineTagRegex.Replace(rawText, match =>
        {
            var expression = match.Groups[1].Value.Trim();
            if (expression.StartsWith(":", StringComparison.Ordinal))
            {
                return string.Empty;
            }

            if (ControlMarker.IsControlToken(expression))
            {
                return string.Empty;
            }

            return ExpressionEvaluator.ToText(ExpressionEvaluator.Evaluate(expression, context));
        });

        if (string.Equals(rawText, replaced, StringComparison.Ordinal))
        {
            return;
        }

        SetParagraphText(paragraph, replaced);
    }

    private static string GetParagraphText(A.Paragraph paragraph)
    {
        return string.Concat(paragraph.Descendants<A.Text>().Select(static text => text.Text));
    }

    private static void SetParagraphText(A.Paragraph paragraph, string text)
    {
        var firstRun = paragraph.Elements<A.Run>().FirstOrDefault();
        var firstRunProperties = firstRun?.RunProperties?.CloneNode(true) as A.RunProperties;
        var removable = paragraph.ChildElements
            .Where(static child => child is A.Run || child is A.Field || child is A.Break)
            .ToList();

        foreach (var child in removable)
        {
            child.Remove();
        }

        if (string.IsNullOrEmpty(text))
        {
            return;
        }

        var run = new A.Run();
        if (firstRunProperties != null)
        {
            run.RunProperties = firstRunProperties;
        }

        run.Text = new A.Text(text);

        var endParagraphRunProperties = paragraph.GetFirstChild<A.EndParagraphRunProperties>();
        if (endParagraphRunProperties != null)
        {
            paragraph.InsertBefore(run, endParagraphRunProperties);
            return;
        }

        paragraph.Append(run);
    }
}

internal static class PptxSlideTagParser
{
    public static bool TryParse(P.Slide slide, out string expression)
    {
        var expressions = slide.Descendants<A.Paragraph>()
            .Select(static paragraph => string.Concat(paragraph.Descendants<A.Text>().Select(static text => text.Text)).Trim())
            .Select(static paragraphText => TagPatterns.SingleTagRegex.Match(paragraphText))
            .Where(static match => match.Success)
            .Select(static match => match.Groups[1].Value.Trim())
            .Where(static token => token.StartsWith(":", StringComparison.Ordinal))
            .Select(static token => token.Substring(1).Trim())
            .Where(static token => token.Length > 0)
            .Distinct(StringComparer.Ordinal)
            .ToList();

        if (expressions.Count == 0)
        {
            expression = string.Empty;
            return false;
        }

        if (expressions.Count > 1)
        {
            throw new InvalidOperationException("Only one slide-level tag is supported per slide in PPTX templates.");
        }

        expression = expressions[0];
        return true;
    }
}

internal static class PptxSlideCloner
{
    public static SlidePart Clone(PresentationPart presentationPart, SlidePart sourceSlidePart)
    {
        var clone = presentationPart.AddNewPart<SlidePart>();
        clone.Slide = (P.Slide)sourceSlidePart.Slide.CloneNode(true);

        foreach (var part in sourceSlidePart.Parts)
        {
            if (part.OpenXmlPart is NotesSlidePart)
            {
                continue;
            }

            clone.AddPart(part.OpenXmlPart, part.RelationshipId);
        }

        foreach (var externalRelationship in sourceSlidePart.ExternalRelationships)
        {
            clone.AddExternalRelationship(
                externalRelationship.RelationshipType,
                externalRelationship.Uri,
                externalRelationship.Id);
        }

        clone.Slide.Save();
        return clone;
    }
}

internal readonly struct PptxSourceSlide
{
    public PptxSourceSlide(P.SlideId slideId, SlidePart slidePart)
    {
        SlideId = slideId;
        SlidePart = slidePart;
    }

    public P.SlideId SlideId { get; }

    public SlidePart SlidePart { get; }
}
