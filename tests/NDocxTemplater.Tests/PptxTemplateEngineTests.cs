using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace NDocxTemplater.Tests;

public class PptxTemplateEngineTests
{
    private readonly PptxTemplateEngine _engine = new PptxTemplateEngine();

    [Fact]
    public void Render_RepeatsTaggedSlides_AndReplacesSlideText()
    {
        var template = CreatePresentation(
            SlideSpec.Create("Title {report.title}"),
            SlideSpec.Create("{:users}", "User {name}", "Amount {amount|format:number:0.00}"),
            SlideSpec.Create("{:showSummary}", "Summary {summaryText}"));

        const string json = @"{
  ""report"": { ""title"": ""Quarterly Review"" },
  ""users"": [
    { ""name"": ""Alice"", ""amount"": 12.5 },
    { ""name"": ""Bob"", ""amount"": 99 }
  ],
  ""showSummary"": true,
  ""summaryText"": ""All regions green""
}";

        var output = _engine.Render(template, json);
        var slides = ReadSlideTexts(output);

        Assert.Equal(4, slides.Count);
        Assert.Equal(new[] { "Title Quarterly Review" }, slides[0]);
        Assert.Equal(new[] { string.Empty, "User Alice", "Amount 12.50" }, slides[1]);
        Assert.Equal(new[] { string.Empty, "User Bob", "Amount 99.00" }, slides[2]);
        Assert.Equal(new[] { string.Empty, "Summary All regions green" }, slides[3]);
    }

    [Fact]
    public void Render_RemovesFalsyTaggedSlides()
    {
        var template = CreatePresentation(
            SlideSpec.Create("Cover {report.title}"),
            SlideSpec.Create("{:users}", "User {name}"),
            SlideSpec.Create("{:showSummary}", "Summary {summaryText}"));

        const string json = @"{
  ""report"": { ""title"": ""May Report"" },
  ""users"": [],
  ""showSummary"": false,
  ""summaryText"": ""should not appear""
}";

        var output = _engine.Render(template, json);
        var slides = ReadSlideTexts(output);

        Assert.Single(slides);
        Assert.Equal(new[] { "Cover May Report" }, slides[0]);
    }

    private static byte[] CreatePresentation(params SlideSpec[] slides)
    {
        using (var stream = new MemoryStream())
        {
            using (var document = PresentationDocument.Create(stream, DocumentFormat.OpenXml.PresentationDocumentType.Presentation, true))
            {
                var presentationPart = document.AddPresentationPart();
                presentationPart.Presentation = new P.Presentation();

                var slideMasterPart = presentationPart.AddNewPart<SlideMasterPart>("rId1");
                var themePart = slideMasterPart.AddNewPart<ThemePart>("rId5");
                themePart.Theme = CreateTheme();
                themePart.Theme.Save();

                var slideLayoutPart = slideMasterPart.AddNewPart<SlideLayoutPart>("rId1");
                slideLayoutPart.SlideLayout = new P.SlideLayout(
                    new P.CommonSlideData(CreateShapeTree()),
                    new P.ColorMapOverride(new A.MasterColorMapping()));
                slideLayoutPart.SlideLayout.Save();

                slideMasterPart.SlideMaster = new P.SlideMaster(
                    new P.CommonSlideData(CreateShapeTree()),
                    new P.ColorMap
                    {
                        Background1 = A.ColorSchemeIndexValues.Light1,
                        Text1 = A.ColorSchemeIndexValues.Dark1,
                        Background2 = A.ColorSchemeIndexValues.Light2,
                        Text2 = A.ColorSchemeIndexValues.Dark2,
                        Accent1 = A.ColorSchemeIndexValues.Accent1,
                        Accent2 = A.ColorSchemeIndexValues.Accent2,
                        Accent3 = A.ColorSchemeIndexValues.Accent3,
                        Accent4 = A.ColorSchemeIndexValues.Accent4,
                        Accent5 = A.ColorSchemeIndexValues.Accent5,
                        Accent6 = A.ColorSchemeIndexValues.Accent6,
                        Hyperlink = A.ColorSchemeIndexValues.Hyperlink,
                        FollowedHyperlink = A.ColorSchemeIndexValues.FollowedHyperlink
                    },
                    new P.SlideLayoutIdList(new P.SlideLayoutId { Id = 2147483649U, RelationshipId = "rId1" }),
                    new P.TextStyles(new P.TitleStyle(), new P.BodyStyle(), new P.OtherStyle()));
                slideMasterPart.SlideMaster.Save();

                presentationPart.Presentation.SlideMasterIdList = new P.SlideMasterIdList(
                    new P.SlideMasterId { Id = 2147483648U, RelationshipId = "rId1" });

                var slideIdList = new P.SlideIdList();
                for (var index = 0; index < slides.Length; index++)
                {
                    var relationshipId = "rIdSlide" + (index + 1).ToString(CultureInfo.InvariantCulture);
                    var slidePart = presentationPart.AddNewPart<SlidePart>(relationshipId);
                    slidePart.AddPart(slideLayoutPart, "rId1");
                    slidePart.Slide = new P.Slide(
                        new P.CommonSlideData(CreateShapeTree(slides[index].Texts)),
                        new P.ColorMapOverride(new A.MasterColorMapping()));
                    slidePart.Slide.Save();

                    slideIdList.Append(new P.SlideId
                    {
                        Id = (uint)(256 + index),
                        RelationshipId = relationshipId
                    });
                }

                presentationPart.Presentation.SlideIdList = slideIdList;
                presentationPart.Presentation.SlideSize = new P.SlideSize { Cx = 9144000, Cy = 6858000 };
                presentationPart.Presentation.NotesSize = new P.NotesSize { Cx = 6858000, Cy = 9144000 };
                presentationPart.Presentation.Save();
            }

            return stream.ToArray();
        }
    }

    private static IReadOnlyList<string[]> ReadSlideTexts(byte[] presentationBytes)
    {
        using (var stream = new MemoryStream(presentationBytes))
        using (var document = PresentationDocument.Open(stream, false))
        {
            return document.PresentationPart!.Presentation.SlideIdList!.Elements<P.SlideId>()
                .Select(slideId =>
                {
                    var slidePart = (SlidePart)document.PresentationPart.GetPartById(slideId.RelationshipId!);
                    return slidePart.Slide.Descendants<A.Paragraph>()
                        .Select(paragraph => string.Concat(paragraph.Descendants<A.Text>().Select(static text => text.Text)))
                        .ToArray();
                })
                .ToArray();
        }
    }

    private static P.ShapeTree CreateShapeTree(params string[] texts)
    {
        var shapeTree = new P.ShapeTree(
            new P.NonVisualGroupShapeProperties(
                new P.NonVisualDrawingProperties { Id = 1U, Name = string.Empty },
                new P.NonVisualGroupShapeDrawingProperties(),
                new P.ApplicationNonVisualDrawingProperties()),
            new P.GroupShapeProperties(new A.TransformGroup()));

        for (var index = 0; index < texts.Length; index++)
        {
            shapeTree.Append(CreateTextShape((uint)(index + 2), "TextBox " + (index + 1).ToString(CultureInfo.InvariantCulture), texts[index], index));
        }

        return shapeTree;
    }

    private static P.Shape CreateTextShape(uint id, string name, string text, int order)
    {
        var offsetY = 400000L + (order * 900000L);
        return new P.Shape(
            new P.NonVisualShapeProperties(
                new P.NonVisualDrawingProperties { Id = id, Name = name },
                new P.NonVisualShapeDrawingProperties(new A.ShapeLocks { NoGrouping = true }),
                new P.ApplicationNonVisualDrawingProperties()),
            new P.ShapeProperties(
                new A.Transform2D(
                    new A.Offset { X = 400000L, Y = offsetY },
                    new A.Extents { Cx = 7600000L, Cy = 700000L }),
                new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }),
            new P.TextBody(
                new A.BodyProperties(),
                new A.ListStyle(),
                new A.Paragraph(
                    new A.Run(new A.Text(text)),
                    new A.EndParagraphRunProperties())));
    }

    private static A.Theme CreateTheme()
    {
        var theme = new A.Theme { Name = "NDocxTemplater Theme" };
        theme.Append(new A.ThemeElements(
                new A.ColorScheme(
                    new A.Dark1Color(new A.SystemColor { Val = A.SystemColorValues.WindowText, LastColor = "000000" }),
                    new A.Light1Color(new A.SystemColor { Val = A.SystemColorValues.Window, LastColor = "FFFFFF" }),
                    new A.Dark2Color(new A.RgbColorModelHex { Val = "1F497D" }),
                    new A.Light2Color(new A.RgbColorModelHex { Val = "EEECE1" }),
                    new A.Accent1Color(new A.RgbColorModelHex { Val = "4F81BD" }),
                    new A.Accent2Color(new A.RgbColorModelHex { Val = "C0504D" }),
                    new A.Accent3Color(new A.RgbColorModelHex { Val = "9BBB59" }),
                    new A.Accent4Color(new A.RgbColorModelHex { Val = "8064A2" }),
                    new A.Accent5Color(new A.RgbColorModelHex { Val = "4BACC6" }),
                    new A.Accent6Color(new A.RgbColorModelHex { Val = "F79646" }),
                    new A.Hyperlink(new A.RgbColorModelHex { Val = "0000FF" }),
                    new A.FollowedHyperlinkColor(new A.RgbColorModelHex { Val = "800080" }))
                { Name = "Office" },
                new A.FontScheme(
                    new A.MajorFont(
                        new A.LatinFont { Typeface = "Calibri" },
                        new A.EastAsianFont { Typeface = string.Empty },
                        new A.ComplexScriptFont { Typeface = string.Empty }),
                    new A.MinorFont(
                        new A.LatinFont { Typeface = "Calibri" },
                        new A.EastAsianFont { Typeface = string.Empty },
                        new A.ComplexScriptFont { Typeface = string.Empty }))
                { Name = "Office" },
                new A.FormatScheme(
                    new A.FillStyleList(
                        new A.SolidFill(new A.SchemeColor { Val = A.SchemeColorValues.PhColor }),
                        new A.GradientFill(
                            new A.GradientStopList(
                                new A.GradientStop(new A.SchemeColor(new A.Tint { Val = 50000 }, new A.SaturationModulation { Val = 300000 }) { Val = A.SchemeColorValues.PhColor }) { Position = 0 },
                                new A.GradientStop(new A.SchemeColor(new A.Tint { Val = 37000 }, new A.SaturationModulation { Val = 300000 }) { Val = A.SchemeColorValues.PhColor }) { Position = 35000 },
                                new A.GradientStop(new A.SchemeColor(new A.Tint { Val = 15000 }, new A.SaturationModulation { Val = 350000 }) { Val = A.SchemeColorValues.PhColor }) { Position = 100000 }),
                            new A.LinearGradientFill { Angle = 16200000, Scaled = true }),
                        new A.GradientFill(
                            new A.GradientStopList(
                                new A.GradientStop(new A.SchemeColor(new A.Shade { Val = 51000 }, new A.SaturationModulation { Val = 130000 }) { Val = A.SchemeColorValues.PhColor }) { Position = 0 },
                                new A.GradientStop(new A.SchemeColor(new A.Shade { Val = 93000 }, new A.SaturationModulation { Val = 130000 }) { Val = A.SchemeColorValues.PhColor }) { Position = 80000 },
                                new A.GradientStop(new A.SchemeColor(new A.Shade { Val = 94000 }, new A.SaturationModulation { Val = 135000 }) { Val = A.SchemeColorValues.PhColor }) { Position = 100000 }),
                            new A.LinearGradientFill { Angle = 16200000, Scaled = false })),
                    new A.LineStyleList(
                        new A.Outline(
                            new A.SolidFill(new A.SchemeColor { Val = A.SchemeColorValues.PhColor }),
                            new A.PresetDash { Val = A.PresetLineDashValues.Solid })
                        { Width = 9525, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center },
                        new A.Outline(
                            new A.SolidFill(new A.SchemeColor { Val = A.SchemeColorValues.PhColor }),
                            new A.PresetDash { Val = A.PresetLineDashValues.Solid })
                        { Width = 25400, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center },
                        new A.Outline(
                            new A.SolidFill(new A.SchemeColor { Val = A.SchemeColorValues.PhColor }),
                            new A.PresetDash { Val = A.PresetLineDashValues.Solid })
                        { Width = 38100, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center }),
                    new A.EffectStyleList(
                        new A.EffectStyle(new A.EffectList()),
                        new A.EffectStyle(new A.EffectList()),
                        new A.EffectStyle(
                            new A.EffectList(
                                new A.OuterShadow
                                {
                                    BlurRadius = 40000,
                                    Distance = 20000,
                                    Direction = 5400000,
                                    RotateWithShape = false
                                }))),
                    new A.BackgroundFillStyleList(
                        new A.SolidFill(new A.SchemeColor { Val = A.SchemeColorValues.PhColor }),
                        new A.GradientFill(
                            new A.GradientStopList(
                                new A.GradientStop(new A.SchemeColor(new A.Tint { Val = 40000 }, new A.SaturationModulation { Val = 350000 }) { Val = A.SchemeColorValues.PhColor }) { Position = 0 },
                                new A.GradientStop(new A.SchemeColor(new A.Tint { Val = 45000 }, new A.Shade { Val = 99000 }, new A.SaturationModulation { Val = 350000 }) { Val = A.SchemeColorValues.PhColor }) { Position = 40000 },
                                new A.GradientStop(new A.SchemeColor(new A.Shade { Val = 20000 }, new A.SaturationModulation { Val = 255000 }) { Val = A.SchemeColorValues.PhColor }) { Position = 100000 }),
                            new A.PathGradientFill { Path = A.PathShadeValues.Circle }),
                        new A.GradientFill(
                            new A.GradientStopList(
                                new A.GradientStop(new A.SchemeColor(new A.Tint { Val = 80000 }, new A.SaturationModulation { Val = 300000 }) { Val = A.SchemeColorValues.PhColor }) { Position = 0 },
                                new A.GradientStop(new A.SchemeColor(new A.Shade { Val = 30000 }, new A.SaturationModulation { Val = 200000 }) { Val = A.SchemeColorValues.PhColor }) { Position = 100000 }),
                            new A.PathGradientFill { Path = A.PathShadeValues.Circle })))
                { Name = "Office" }));
        return theme;
    }

    private readonly struct SlideSpec
    {
        public SlideSpec(string[] texts)
        {
            Texts = texts;
        }

        public string[] Texts { get; }

        public static SlideSpec Create(params string[] texts)
        {
            return new SlideSpec(texts);
        }
    }
}
