using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.Json.Nodes;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
using JArray = System.Text.Json.Nodes.JsonArray;
using JObject = System.Text.Json.Nodes.JsonObject;
using JToken = System.Text.Json.Nodes.JsonNode;
using SixLabors.ImageSharp;
using SixLabors.ImageSharp.PixelFormats;
using ZXing;
using ZXing.Common;

namespace NDocxTemplater;

internal static class ImageTemplateRenderer
{
    public static bool TryRenderImageTag(
        Paragraph paragraph,
        TemplateContext context,
        MainDocumentPart mainDocumentPart,
        Func<uint> nextImageId)
    {
        if (!ImageTagParser.TryParseParagraph(paragraph, out var imageTag))
        {
            return false;
        }

        var images = ResolvePayloads(imageTag, context).ToList();

        foreach (var run in paragraph.Elements<Run>().ToList())
        {
            run.Remove();
        }

        if (imageTag.Centered)
        {
            CenterParagraph(paragraph);
        }

        foreach (var image in images)
        {
            paragraph.Append(CreateImageRun(mainDocumentPart, image, nextImageId()));
        }

        return true;
    }

    private static IEnumerable<ImagePayload> ResolvePayloads(ImageTag imageTag, TemplateContext context)
    {
        if (BarcodeTemplateParser.IsBarcodeExpression(imageTag.Expression))
        {
            var barcodeTemplate = BarcodeTemplateParser.Parse(imageTag.Expression);
            return BarcodeTemplateRenderer.ResolveMany(barcodeTemplate, context);
        }

        var imageToken = ExpressionEvaluator.Evaluate(imageTag.Expression, context);
        return ImageInputResolver.ResolveMany(imageToken);
    }

    private static void CenterParagraph(Paragraph paragraph)
    {
        var properties = paragraph.GetFirstChild<ParagraphProperties>();
        if (properties == null)
        {
            properties = paragraph.PrependChild(new ParagraphProperties());
        }

        properties.Justification = new Justification { Val = JustificationValues.Center };
    }

    private static Run CreateImageRun(MainDocumentPart mainDocumentPart, ImagePayload image, uint imageId)
    {
        var imagePart = mainDocumentPart.AddImagePart(image.ImagePartType);
        using (var imageStream = new MemoryStream(image.Bytes, writable: false))
        {
            imagePart.FeedData(imageStream);
        }

        var relationId = mainDocumentPart.GetIdOfPart(imagePart);
        var drawing = CreateDrawing(relationId, image.WidthPx, image.HeightPx, imageId);
        return new Run(drawing);
    }

    private static Drawing CreateDrawing(string relationId, int widthPx, int heightPx, uint imageId)
    {
        var widthEmu = PixelsToEmu(widthPx);
        var heightEmu = PixelsToEmu(heightPx);

        return new Drawing(
            new DW.Inline(
                new DW.Extent { Cx = widthEmu, Cy = heightEmu },
                new DW.EffectExtent
                {
                    LeftEdge = 0L,
                    TopEdge = 0L,
                    RightEdge = 0L,
                    BottomEdge = 0L
                },
                new DW.DocProperties { Id = imageId, Name = "Image " + imageId.ToString(CultureInfo.InvariantCulture) },
                new DW.NonVisualGraphicFrameDrawingProperties(
                    new A.GraphicFrameLocks { NoChangeAspect = true }),
                new A.Graphic(
                    new A.GraphicData(
                        new PIC.Picture(
                            new PIC.NonVisualPictureProperties(
                                new PIC.NonVisualDrawingProperties { Id = imageId, Name = "Image " + imageId.ToString(CultureInfo.InvariantCulture) },
                                new PIC.NonVisualPictureDrawingProperties()),
                            new PIC.BlipFill(
                                new A.Blip
                                {
                                    Embed = relationId,
                                    CompressionState = A.BlipCompressionValues.Print
                                },
                                new A.Stretch(new A.FillRectangle())),
                            new PIC.ShapeProperties(
                                new A.Transform2D(
                                    new A.Offset { X = 0L, Y = 0L },
                                    new A.Extents { Cx = widthEmu, Cy = heightEmu }),
                                new A.PresetGeometry(new A.AdjustValueList())
                                {
                                    Preset = A.ShapeTypeValues.Rectangle
                                })))
                    {
                        Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture"
                    }))
            {
                DistanceFromTop = 0U,
                DistanceFromBottom = 0U,
                DistanceFromLeft = 0U,
                DistanceFromRight = 0U
            });
    }

    private static long PixelsToEmu(int pixels)
    {
        var safePixels = pixels <= 0 ? 1 : pixels;
        return safePixels * 9525L;
    }
}

internal static class ImageTagParser
{
    public static bool TryParseParagraph(Paragraph paragraph, out ImageTag imageTag)
    {
        imageTag = default;

        var rawText = string.Concat(paragraph.Descendants<Text>().Select(static text => text.Text)).Trim();
        if (rawText.Length == 0)
        {
            return false;
        }

        var fullTag = TagPatterns.SingleTagRegex.Match(rawText);
        if (!fullTag.Success)
        {
            return false;
        }

        var token = fullTag.Groups[1].Value.Trim();
        return TryParseToken(token, out imageTag);
    }

    public static bool TryParseToken(string token, out ImageTag imageTag)
    {
        imageTag = default;

        if (string.IsNullOrWhiteSpace(token) || !token.StartsWith("%", StringComparison.Ordinal))
        {
            return false;
        }

        var centered = token.StartsWith("%%", StringComparison.Ordinal);
        var expression = centered
            ? token.Substring(2).Trim()
            : token.Substring(1).Trim();

        if (expression.Length == 0)
        {
            return false;
        }

        imageTag = new ImageTag(expression, centered);
        return true;
    }
}

internal static class ImageInputResolver
{
    public static IEnumerable<ImagePayload> ResolveMany(JToken? token)
    {
        if (JsonNodeHelpers.IsNull(token))
        {
            return Enumerable.Empty<ImagePayload>();
        }

        if (token is JArray array)
        {
            return array.Where(static item => item != null).Select(static item => ResolveSingle(item!)).ToList();
        }

        return new[] { ResolveSingle(token!) };
    }

    private static ImagePayload ResolveSingle(JToken token)
    {
        string? source = null;
        int? width = null;
        int? height = null;
        int? maxWidth = null;
        int? maxHeight = null;
        double? scale = null;
        bool? preserveAspectRatio = null;

        if (JsonNodeHelpers.TryGetString(token, out var stringToken))
        {
            source = stringToken;
        }
        else if (token is JObject obj)
        {
            source = ReadString(obj, "src")
                ?? ReadString(obj, "data")
                ?? ReadString(obj, "base64")
                ?? ReadString(obj, "path")
                ?? ReadString(obj, "value");

            width = ReadInteger(obj, "width") ?? ReadInteger(obj, "widthPx");
            height = ReadInteger(obj, "height") ?? ReadInteger(obj, "heightPx");
            maxWidth = ReadInteger(obj, "maxWidth") ?? ReadInteger(obj, "maxWidthPx");
            maxHeight = ReadInteger(obj, "maxHeight") ?? ReadInteger(obj, "maxHeightPx");
            scale = ReadDouble(obj, "scale") ?? ReadDouble(obj, "scaleRatio");
            preserveAspectRatio = ReadBoolean(obj, "preserveAspectRatio")
                ?? ReadBoolean(obj, "keepAspectRatio")
                ?? ReadBoolean(obj, "lockAspectRatio");
        }

        var sourceText = source?.Trim();
        if (string.IsNullOrWhiteSpace(sourceText))
        {
            throw new InvalidOperationException("Image value must be a string or object containing src/data/base64/path.");
        }

        var imageBytes = ParseImageBytes(sourceText!, out var mimeHint, out var extensionHint);
        var imagePartType = DetectImagePartType(imageBytes, mimeHint, extensionHint);

        var inferredSize = ImageBinaryInspector.TryReadPixelSize(imageBytes);
        var resolvedSize = ResolveOutputSize(
            inferredSize,
            width,
            height,
            maxWidth,
            maxHeight,
            scale,
            preserveAspectRatio);

        return new ImagePayload(imageBytes, imagePartType, resolvedSize.Width, resolvedSize.Height);
    }

    private static byte[] ParseImageBytes(string source, out string? mimeHint, out string? extensionHint)
    {
        mimeHint = null;
        extensionHint = null;

        if (source.StartsWith("data:", StringComparison.OrdinalIgnoreCase))
        {
            var commaIndex = source.IndexOf(',');
            if (commaIndex < 0)
            {
                throw new InvalidOperationException("Invalid data URI for image value.");
            }

            var header = source.Substring(5, commaIndex - 5);
            var payload = source.Substring(commaIndex + 1);

            var segments = header.Split(';');
            if (segments.Length > 0)
            {
                mimeHint = segments[0];
            }

            var isBase64 = segments.Any(static item => string.Equals(item, "base64", StringComparison.OrdinalIgnoreCase));
            if (!isBase64)
            {
                throw new InvalidOperationException("Only base64 data URI image values are supported.");
            }

            return Convert.FromBase64String(payload);
        }

        if (File.Exists(source))
        {
            extensionHint = Path.GetExtension(source);
            return File.ReadAllBytes(source);
        }

        try
        {
            return Convert.FromBase64String(source);
        }
        catch (System.FormatException)
        {
            throw new InvalidOperationException(
                "Image string value must be base64, base64 data URI, or an existing file path.");
        }
    }

    private static PartTypeInfo DetectImagePartType(byte[] bytes, string? mimeHint, string? extensionHint)
    {
        var mimeType = mimeHint?.Trim().ToLowerInvariant();
        if (!string.IsNullOrEmpty(mimeType))
        {
            switch (mimeType)
            {
                case "image/png":
                    return ImagePartType.Png;
                case "image/jpeg":
                case "image/jpg":
                    return ImagePartType.Jpeg;
                case "image/gif":
                    return ImagePartType.Gif;
                case "image/bmp":
                    return ImagePartType.Bmp;
                case "image/tiff":
                    return ImagePartType.Tiff;
            }
        }

        if (ImageBinaryInspector.IsPng(bytes))
        {
            return ImagePartType.Png;
        }

        if (ImageBinaryInspector.IsJpeg(bytes))
        {
            return ImagePartType.Jpeg;
        }

        if (ImageBinaryInspector.IsGif(bytes))
        {
            return ImagePartType.Gif;
        }

        if (ImageBinaryInspector.IsBmp(bytes))
        {
            return ImagePartType.Bmp;
        }

        if (ImageBinaryInspector.IsTiff(bytes))
        {
            return ImagePartType.Tiff;
        }

        var ext = extensionHint?.Trim().ToLowerInvariant();
        if (!string.IsNullOrEmpty(ext))
        {
            switch (ext)
            {
                case ".png":
                    return ImagePartType.Png;
                case ".jpg":
                case ".jpeg":
                    return ImagePartType.Jpeg;
                case ".gif":
                    return ImagePartType.Gif;
                case ".bmp":
                    return ImagePartType.Bmp;
                case ".tif":
                case ".tiff":
                    return ImagePartType.Tiff;
            }
        }

        throw new InvalidOperationException("Unable to detect image type. Supported types: png, jpeg, gif, bmp, tiff.");
    }

    private static string? ReadString(JObject obj, string key)
    {
        foreach (var pair in obj)
        {
            if (!string.Equals(pair.Key, key, StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            if (JsonNodeHelpers.TryGetString(pair.Value, out var text))
            {
                return text;
            }

            return pair.Value?.ToJsonString();
        }

        return null;
    }

    private static int? ReadInteger(JObject obj, string key)
    {
        foreach (var pair in obj)
        {
            if (!string.Equals(pair.Key, key, StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            var token = pair.Value;
            if (JsonNodeHelpers.IsNull(token))
            {
                return null;
            }

            if (token is JsonValue jsonValue)
            {
                try
                {
                    return jsonValue.GetValue<int>();
                }
                catch
                {
                }
            }

            if (int.TryParse(token!.ToJsonString().Trim('\"'), NumberStyles.Integer, CultureInfo.InvariantCulture, out var value))
            {
                return value;
            }

            return null;
        }

        return null;
    }

    private static double? ReadDouble(JObject obj, string key)
    {
        foreach (var pair in obj)
        {
            if (!string.Equals(pair.Key, key, StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            var token = pair.Value;
            if (JsonNodeHelpers.IsNull(token))
            {
                return null;
            }

            if (token is JsonValue jsonValue)
            {
                try
                {
                    return jsonValue.GetValue<double>();
                }
                catch
                {
                }
            }

            if (double.TryParse(token!.ToJsonString().Trim('\"'), NumberStyles.Float, CultureInfo.InvariantCulture, out var value))
            {
                return value;
            }

            return null;
        }

        return null;
    }

    private static bool? ReadBoolean(JObject obj, string key)
    {
        foreach (var pair in obj)
        {
            if (!string.Equals(pair.Key, key, StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            var token = pair.Value;
            if (JsonNodeHelpers.IsNull(token))
            {
                return null;
            }

            if (token is JsonValue jsonValue)
            {
                try
                {
                    return jsonValue.GetValue<bool>();
                }
                catch
                {
                }
            }

            if (bool.TryParse(token!.ToJsonString().Trim('\"'), out var value))
            {
                return value;
            }

            return null;
        }

        return null;
    }

    private static ImageSize ResolveOutputSize(
        ImageSize? inferredSize,
        int? width,
        int? height,
        int? maxWidth,
        int? maxHeight,
        double? scale,
        bool? preserveAspectRatio)
    {
        if (width.HasValue && width.Value <= 0)
        {
            throw new InvalidOperationException("Image width must be greater than zero.");
        }

        if (height.HasValue && height.Value <= 0)
        {
            throw new InvalidOperationException("Image height must be greater than zero.");
        }

        if (maxWidth.HasValue && maxWidth.Value <= 0)
        {
            throw new InvalidOperationException("Image maxWidth must be greater than zero.");
        }

        if (maxHeight.HasValue && maxHeight.Value <= 0)
        {
            throw new InvalidOperationException("Image maxHeight must be greater than zero.");
        }

        if (scale.HasValue && scale.Value <= 0d)
        {
            throw new InvalidOperationException("Image scale must be greater than zero.");
        }

        var hasScaleConstraints = scale.HasValue || maxWidth.HasValue || maxHeight.HasValue;
        var hasOneDimensionOnly = width.HasValue ^ height.HasValue;
        var keepAspect = preserveAspectRatio ?? hasScaleConstraints || hasOneDimensionOnly;

        var originalWidth = inferredSize?.Width;
        var originalHeight = inferredSize?.Height;

        int targetWidth;
        int targetHeight;

        if (width.HasValue && height.HasValue)
        {
            if (keepAspect && originalWidth.HasValue && originalHeight.HasValue)
            {
                (targetWidth, targetHeight) = FitIntoBox(originalWidth.Value, originalHeight.Value, width.Value, height.Value, allowUpscale: true);
            }
            else
            {
                targetWidth = width.Value;
                targetHeight = height.Value;
            }
        }
        else if (width.HasValue)
        {
            targetWidth = width.Value;
            if (keepAspect && originalWidth.HasValue && originalHeight.HasValue)
            {
                targetHeight = ScaleDimension(originalHeight.Value, width.Value / (double)originalWidth.Value);
            }
            else
            {
                targetHeight = originalHeight ?? 120;
            }
        }
        else if (height.HasValue)
        {
            targetHeight = height.Value;
            if (keepAspect && originalWidth.HasValue && originalHeight.HasValue)
            {
                targetWidth = ScaleDimension(originalWidth.Value, height.Value / (double)originalHeight.Value);
            }
            else
            {
                targetWidth = originalWidth ?? 120;
            }
        }
        else if (originalWidth.HasValue && originalHeight.HasValue)
        {
            targetWidth = originalWidth.Value;
            targetHeight = originalHeight.Value;
        }
        else
        {
            targetWidth = 120;
            targetHeight = 120;
        }

        if (scale.HasValue)
        {
            targetWidth = ScaleDimension(targetWidth, scale.Value);
            targetHeight = ScaleDimension(targetHeight, scale.Value);
        }

        if (maxWidth.HasValue || maxHeight.HasValue)
        {
            if (keepAspect)
            {
                var fitWidth = maxWidth ?? int.MaxValue;
                var fitHeight = maxHeight ?? int.MaxValue;
                (targetWidth, targetHeight) = FitIntoBox(targetWidth, targetHeight, fitWidth, fitHeight, allowUpscale: false);
            }
            else
            {
                if (maxWidth.HasValue && targetWidth > maxWidth.Value)
                {
                    targetWidth = maxWidth.Value;
                }

                if (maxHeight.HasValue && targetHeight > maxHeight.Value)
                {
                    targetHeight = maxHeight.Value;
                }
            }
        }

        if (targetWidth <= 0 || targetHeight <= 0)
        {
            throw new InvalidOperationException("Resolved image width and height must be greater than zero.");
        }

        return new ImageSize(targetWidth, targetHeight);
    }

    private static (int Width, int Height) FitIntoBox(int sourceWidth, int sourceHeight, int boxWidth, int boxHeight, bool allowUpscale)
    {
        if (sourceWidth <= 0 || sourceHeight <= 0)
        {
            throw new InvalidOperationException("Source image dimensions must be greater than zero.");
        }

        if (boxWidth <= 0 || boxHeight <= 0)
        {
            throw new InvalidOperationException("Image fit box dimensions must be greater than zero.");
        }

        var ratioX = boxWidth / (double)sourceWidth;
        var ratioY = boxHeight / (double)sourceHeight;
        var ratio = Math.Min(ratioX, ratioY);
        if (!allowUpscale)
        {
            ratio = Math.Min(ratio, 1d);
        }

        return (ScaleDimension(sourceWidth, ratio), ScaleDimension(sourceHeight, ratio));
    }

    private static int ScaleDimension(int value, double scale)
    {
        if (value <= 0)
        {
            throw new InvalidOperationException("Image dimension must be greater than zero.");
        }

        if (scale <= 0d)
        {
            throw new InvalidOperationException("Image scale factor must be greater than zero.");
        }

        var scaled = (int)Math.Round(value * scale, MidpointRounding.AwayFromZero);
        return scaled <= 0 ? 1 : scaled;
    }
}

internal static class BarcodeTemplateParser
{
    public static bool IsBarcodeExpression(string expression)
    {
        return !string.IsNullOrWhiteSpace(expression)
            && expression.TrimStart().StartsWith("barcode:", StringComparison.OrdinalIgnoreCase);
    }

    public static BarcodeTemplate Parse(string expression)
    {
        if (string.IsNullOrWhiteSpace(expression))
        {
            throw new InvalidOperationException("Barcode expression cannot be empty.");
        }

        var trimmed = expression.Trim();
        if (!trimmed.StartsWith("barcode:", StringComparison.OrdinalIgnoreCase))
        {
            throw new InvalidOperationException("Barcode expression must start with 'barcode:'.");
        }

        var body = trimmed.Substring("barcode:".Length).Trim();
        if (body.Length == 0)
        {
            throw new InvalidOperationException("Barcode expression must include a value path.");
        }

        var segments = body.Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries)
            .Select(static part => part.Trim())
            .Where(static part => part.Length > 0)
            .ToArray();

        if (segments.Length == 0)
        {
            throw new InvalidOperationException("Barcode expression must include a value path.");
        }

        var valueExpression = segments[0];
        if (valueExpression.Length == 0)
        {
            throw new InvalidOperationException("Barcode value path cannot be empty.");
        }

        var barcodeType = BarcodeType.Code128;
        int? width = null;
        int? height = null;
        int? margin = null;
        bool? pureBarcode = null;

        foreach (var segment in segments.Skip(1))
        {
            var separatorIndex = segment.IndexOf('=');
            if (separatorIndex < 0)
            {
                throw new InvalidOperationException(
                    string.Format(
                        CultureInfo.InvariantCulture,
                        "Invalid barcode parameter '{0}'. Expected key=value.",
                        segment));
            }

            var key = segment.Substring(0, separatorIndex).Trim();
            var value = segment.Substring(separatorIndex + 1).Trim();

            if (key.Length == 0)
            {
                throw new InvalidOperationException("Barcode parameter key cannot be empty.");
            }

            switch (key.ToLowerInvariant())
            {
                case "type":
                case "format":
                case "barcodetype":
                    barcodeType = ParseBarcodeType(value);
                    break;
                case "width":
                case "widthpx":
                    width = ParsePositiveInteger(value, key);
                    break;
                case "height":
                case "heightpx":
                    height = ParsePositiveInteger(value, key);
                    break;
                case "margin":
                    margin = ParseNonNegativeInteger(value, key);
                    break;
                case "pure":
                case "purebarcode":
                    pureBarcode = ParseBoolean(value, key);
                    break;
                default:
                    throw new InvalidOperationException(
                        string.Format(
                            CultureInfo.InvariantCulture,
                            "Unsupported barcode parameter '{0}'. Supported: type, width, height, margin, pure.",
                            key));
            }
        }

        return new BarcodeTemplate(valueExpression, barcodeType, width, height, margin, pureBarcode);
    }

    private static BarcodeType ParseBarcodeType(string text)
    {
        var normalized = (text ?? string.Empty)
            .Trim()
            .ToLowerInvariant()
            .Replace("-", string.Empty)
            .Replace("_", string.Empty)
            .Replace(" ", string.Empty);

        switch (normalized)
        {
            case "code128":
                return BarcodeType.Code128;
            case "code39":
                return BarcodeType.Code39;
            case "code93":
                return BarcodeType.Code93;
            case "codabar":
                return BarcodeType.Codabar;
            case "ean13":
                return BarcodeType.Ean13;
            case "ean8":
                return BarcodeType.Ean8;
            case "upca":
                return BarcodeType.UpcA;
            case "itf":
            case "interleaved2of5":
                return BarcodeType.Itf;
            default:
                throw new InvalidOperationException(
                    string.Format(
                        CultureInfo.InvariantCulture,
                        "Unsupported barcode type '{0}'. Supported: code128, code39, code93, codabar, ean13, ean8, upca, itf.",
                        text));
        }
    }

    private static int ParsePositiveInteger(string text, string key)
    {
        if (!int.TryParse(text, NumberStyles.Integer, CultureInfo.InvariantCulture, out var value) || value <= 0)
        {
            throw new InvalidOperationException(
                string.Format(
                    CultureInfo.InvariantCulture,
                    "Barcode parameter '{0}' must be a positive integer.",
                    key));
        }

        return value;
    }

    private static int ParseNonNegativeInteger(string text, string key)
    {
        if (!int.TryParse(text, NumberStyles.Integer, CultureInfo.InvariantCulture, out var value) || value < 0)
        {
            throw new InvalidOperationException(
                string.Format(
                    CultureInfo.InvariantCulture,
                    "Barcode parameter '{0}' must be a non-negative integer.",
                    key));
        }

        return value;
    }

    private static bool ParseBoolean(string text, string key)
    {
        if (!bool.TryParse(text, out var value))
        {
            throw new InvalidOperationException(
                string.Format(
                    CultureInfo.InvariantCulture,
                    "Barcode parameter '{0}' must be a boolean value.",
                    key));
        }

        return value;
    }
}

internal static class BarcodeTemplateRenderer
{
    private const int DefaultWidth = 320;
    private const int DefaultHeight = 96;
    private const int DefaultMargin = 2;
    private const bool DefaultPureBarcode = true;

    public static IEnumerable<ImagePayload> ResolveMany(BarcodeTemplate template, TemplateContext context)
    {
        var valueToken = ExpressionEvaluator.Evaluate(template.ValueExpression, context);
        if (JsonNodeHelpers.IsNull(valueToken))
        {
            return Enumerable.Empty<ImagePayload>();
        }

        if (valueToken is JArray array)
        {
            return array
                .Where(static item => !JsonNodeHelpers.IsNull(item))
                .Select(item => CreatePayload(template, item!))
                .ToArray();
        }

        return new[] { CreatePayload(template, valueToken!) };
    }

    private static ImagePayload CreatePayload(BarcodeTemplate template, JToken valueToken)
    {
        var rawValue = ExpressionEvaluator.ToText(valueToken).Trim();
        if (string.IsNullOrWhiteSpace(rawValue))
        {
            throw new InvalidOperationException("Barcode value cannot be empty.");
        }

        var format = MapBarcodeFormat(template.Type);
        var width = template.Width ?? DefaultWidth;
        var height = template.Height ?? DefaultHeight;
        var margin = template.Margin ?? DefaultMargin;
        var pureBarcode = template.PureBarcode ?? DefaultPureBarcode;

        var writer = new BarcodeWriterPixelData
        {
            Format = format,
            Options = new EncodingOptions
            {
                Width = width,
                Height = height,
                Margin = margin,
                PureBarcode = pureBarcode
            }
        };

        var pixelData = writer.Write(rawValue);

        using (var image = Image.LoadPixelData<Bgra32>(pixelData.Pixels, pixelData.Width, pixelData.Height))
        using (var buffer = new MemoryStream())
        {
            image.SaveAsPng(buffer);
            return new ImagePayload(buffer.ToArray(), ImagePartType.Png, pixelData.Width, pixelData.Height);
        }
    }

    private static BarcodeFormat MapBarcodeFormat(BarcodeType type)
    {
        switch (type)
        {
            case BarcodeType.Code128:
                return BarcodeFormat.CODE_128;
            case BarcodeType.Code39:
                return BarcodeFormat.CODE_39;
            case BarcodeType.Code93:
                return BarcodeFormat.CODE_93;
            case BarcodeType.Codabar:
                return BarcodeFormat.CODABAR;
            case BarcodeType.Ean13:
                return BarcodeFormat.EAN_13;
            case BarcodeType.Ean8:
                return BarcodeFormat.EAN_8;
            case BarcodeType.UpcA:
                return BarcodeFormat.UPC_A;
            case BarcodeType.Itf:
                return BarcodeFormat.ITF;
            default:
                throw new InvalidOperationException("Unsupported barcode type.");
        }
    }
}

internal static class ImageBinaryInspector
{
    public static bool IsPng(byte[] data)
    {
        return data.Length >= 8
            && data[0] == 0x89
            && data[1] == 0x50
            && data[2] == 0x4E
            && data[3] == 0x47
            && data[4] == 0x0D
            && data[5] == 0x0A
            && data[6] == 0x1A
            && data[7] == 0x0A;
    }

    public static bool IsJpeg(byte[] data)
    {
        return data.Length >= 3
            && data[0] == 0xFF
            && data[1] == 0xD8
            && data[2] == 0xFF;
    }

    public static bool IsGif(byte[] data)
    {
        if (data.Length < 10)
        {
            return false;
        }

        return data[0] == 0x47
            && data[1] == 0x49
            && data[2] == 0x46
            && data[3] == 0x38
            && (data[4] == 0x37 || data[4] == 0x39)
            && data[5] == 0x61;
    }

    public static bool IsBmp(byte[] data)
    {
        return data.Length >= 2 && data[0] == 0x42 && data[1] == 0x4D;
    }

    public static bool IsTiff(byte[] data)
    {
        if (data.Length < 4)
        {
            return false;
        }

        var littleEndian = data[0] == 0x49 && data[1] == 0x49 && data[2] == 0x2A && data[3] == 0x00;
        var bigEndian = data[0] == 0x4D && data[1] == 0x4D && data[2] == 0x00 && data[3] == 0x2A;
        return littleEndian || bigEndian;
    }

    public static ImageSize? TryReadPixelSize(byte[] data)
    {
        if (TryReadPngSize(data, out var pngSize))
        {
            return pngSize;
        }

        if (TryReadGifSize(data, out var gifSize))
        {
            return gifSize;
        }

        if (TryReadJpegSize(data, out var jpegSize))
        {
            return jpegSize;
        }

        return null;
    }

    private static bool TryReadPngSize(byte[] data, out ImageSize size)
    {
        size = default;
        if (!IsPng(data) || data.Length < 24)
        {
            return false;
        }

        var width = ReadInt32BigEndian(data, 16);
        var height = ReadInt32BigEndian(data, 20);
        if (width <= 0 || height <= 0)
        {
            return false;
        }

        size = new ImageSize(width, height);
        return true;
    }

    private static bool TryReadGifSize(byte[] data, out ImageSize size)
    {
        size = default;
        if (!IsGif(data))
        {
            return false;
        }

        var width = data[6] | (data[7] << 8);
        var height = data[8] | (data[9] << 8);
        if (width <= 0 || height <= 0)
        {
            return false;
        }

        size = new ImageSize(width, height);
        return true;
    }

    private static bool TryReadJpegSize(byte[] data, out ImageSize size)
    {
        size = default;
        if (!IsJpeg(data) || data.Length < 4)
        {
            return false;
        }

        var index = 2;
        while (index + 8 < data.Length)
        {
            if (data[index] != 0xFF)
            {
                index++;
                continue;
            }

            while (index < data.Length && data[index] == 0xFF)
            {
                index++;
            }

            if (index >= data.Length)
            {
                break;
            }

            var marker = data[index++];
            if (marker == 0xD8 || marker == 0xD9)
            {
                continue;
            }

            if (index + 1 >= data.Length)
            {
                break;
            }

            var segmentLength = (data[index] << 8) + data[index + 1];
            if (segmentLength < 2 || index + segmentLength > data.Length)
            {
                break;
            }

            var isStartOfFrame = marker >= 0xC0 && marker <= 0xCF
                && marker != 0xC4
                && marker != 0xC8
                && marker != 0xCC;

            if (isStartOfFrame && segmentLength >= 7)
            {
                var height = (data[index + 3] << 8) + data[index + 4];
                var width = (data[index + 5] << 8) + data[index + 6];
                if (width > 0 && height > 0)
                {
                    size = new ImageSize(width, height);
                    return true;
                }
            }

            index += segmentLength;
        }

        return false;
    }

    private static int ReadInt32BigEndian(byte[] data, int offset)
    {
        return (data[offset] << 24)
            | (data[offset + 1] << 16)
            | (data[offset + 2] << 8)
            | data[offset + 3];
    }
}

internal readonly struct ImageTag
{
    public ImageTag(string expression, bool centered)
    {
        Expression = expression;
        Centered = centered;
    }

    public string Expression { get; }

    public bool Centered { get; }
}

internal readonly struct ImagePayload
{
    public ImagePayload(byte[] bytes, PartTypeInfo imagePartType, int widthPx, int heightPx)
    {
        Bytes = bytes;
        ImagePartType = imagePartType;
        WidthPx = widthPx;
        HeightPx = heightPx;
    }

    public byte[] Bytes { get; }

    public PartTypeInfo ImagePartType { get; }

    public int WidthPx { get; }

    public int HeightPx { get; }
}

internal readonly struct ImageSize
{
    public ImageSize(int width, int height)
    {
        Width = width;
        Height = height;
    }

    public int Width { get; }

    public int Height { get; }
}

internal readonly struct BarcodeTemplate
{
    public BarcodeTemplate(
        string valueExpression,
        BarcodeType type,
        int? width,
        int? height,
        int? margin,
        bool? pureBarcode)
    {
        ValueExpression = valueExpression;
        Type = type;
        Width = width;
        Height = height;
        Margin = margin;
        PureBarcode = pureBarcode;
    }

    public string ValueExpression { get; }

    public BarcodeType Type { get; }

    public int? Width { get; }

    public int? Height { get; }

    public int? Margin { get; }

    public bool? PureBarcode { get; }
}

internal enum BarcodeType
{
    Code128,
    Code39,
    Code93,
    Codabar,
    Ean13,
    Ean8,
    UpcA,
    Itf
}
