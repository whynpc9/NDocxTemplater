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

        var imageToken = ExpressionEvaluator.Evaluate(imageTag.Expression, context);
        var images = ImageInputResolver.ResolveMany(imageToken).ToList();

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
        }

        var sourceText = source?.Trim();
        if (string.IsNullOrWhiteSpace(sourceText))
        {
            throw new InvalidOperationException("Image value must be a string or object containing src/data/base64/path.");
        }

        var imageBytes = ParseImageBytes(sourceText!, out var mimeHint, out var extensionHint);
        var imagePartType = DetectImagePartType(imageBytes, mimeHint, extensionHint);

        var inferredSize = ImageBinaryInspector.TryReadPixelSize(imageBytes);
        var widthPx = width;
        var heightPx = height;

        if (widthPx == null && inferredSize.HasValue)
        {
            widthPx = inferredSize.Value.Width;
        }

        if (heightPx == null && inferredSize.HasValue)
        {
            heightPx = inferredSize.Value.Height;
        }

        if (widthPx == null)
        {
            widthPx = 120;
        }

        if (heightPx == null)
        {
            heightPx = 120;
        }

        if (widthPx.Value <= 0 || heightPx.Value <= 0)
        {
            throw new InvalidOperationException("Image width and height must be greater than zero.");
        }

        return new ImagePayload(imageBytes, imagePartType, widthPx.Value, heightPx.Value);
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
        catch (FormatException)
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
