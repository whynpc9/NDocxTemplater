using System;
using System.IO;
using System.Linq;
using System.IO.Compression;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using NDocxTemplater.Cli;
using Xunit;

namespace NDocxTemplater.Tests;

public class CliApplicationTests
{
    [Fact]
    public void Render_WritesDocxOutput()
    {
        using var temp = new TemporaryDirectory();
        var templatePath = Path.Combine(temp.Path, "template.docx");
        var dataPath = Path.Combine(temp.Path, "data.json");
        var outputPath = Path.Combine(temp.Path, "output.docx");
        File.WriteAllBytes(templatePath, CreateTemplate(Paragraph("Hello {name}")));
        File.WriteAllText(dataPath, @"{ ""name"": ""Alice"" }");

        var exitCode = CliApplication.Run(
            new[] { "render", "--template", templatePath, "--data", dataPath, "--output", outputPath },
            TextWriter.Null,
            TextWriter.Null);

        Assert.Equal(0, exitCode);
        Assert.True(File.Exists(outputPath));
        Assert.Equal("Hello Alice", ReadBodyText(File.ReadAllBytes(outputPath)));
    }

    [Fact]
    public void Validate_ReturnsFailureForMissingTags_ByDefault()
    {
        using var temp = new TemporaryDirectory();
        var templatePath = Path.Combine(temp.Path, "template.docx");
        var dataPath = Path.Combine(temp.Path, "data.json");
        using var error = new StringWriter();
        File.WriteAllBytes(templatePath, CreateTemplate(Paragraph("Hello {missing.name}")));
        File.WriteAllText(dataPath, @"{ ""name"": ""Alice"" }");

        var exitCode = CliApplication.Run(
            new[] { "validate", "--template", templatePath, "--data", dataPath },
            TextWriter.Null,
            error);

        Assert.Equal(1, exitCode);
        Assert.Contains("missing.name", error.ToString(), StringComparison.Ordinal);
    }

    [Fact]
    public void InspectTags_PrintsTemplateTags()
    {
        using var temp = new TemporaryDirectory();
        var templatePath = Path.Combine(temp.Path, "template.docx");
        using var output = new StringWriter();
        File.WriteAllBytes(templatePath, CreateTemplate(Paragraph("Hello {name}"), Paragraph("{#items}"), Paragraph("{value}"), Paragraph("{/items}")));

        var exitCode = CliApplication.Run(
            new[] { "inspect-tags", "--template", templatePath },
            output,
            TextWriter.Null);

        var lines = output.ToString()
            .Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);

        Assert.Equal(0, exitCode);
        Assert.Contains("#items", lines);
        Assert.Contains("/items", lines);
        Assert.Contains("name", lines);
        Assert.Contains("value", lines);
    }

    [Theory]
    [InlineData("docx", "word/document.xml", "http://schemas.openxmlformats.org/wordprocessingml/2006/main", "p", "t")]
    [InlineData("pptx", "ppt/slides/slide1.xml", "http://schemas.openxmlformats.org/drawingml/2006/main", "p", "t")]
    [InlineData("xlsx", "xl/sharedStrings.xml", "http://schemas.openxmlformats.org/spreadsheetml/2006/main", "si", "t")]
    public void InspectTags_ReadsLogicalTextAcrossRuns_WithoutCrossingParagraphs(
        string format, string partPath, string xmlNamespace, string containerName, string textName)
    {
        using var temp = new TemporaryDirectory();
        var templatePath = Path.Combine(temp.Path, "template." + format);
        XNamespace ns = xmlNamespace;
        using (var stream = File.Create(templatePath))
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Create))
        {
            var entry = archive.CreateEntry(partPath);
            using var xmlStream = entry.Open();
            new XDocument(new XElement(ns + "root",
                new XElement(ns + containerName,
                    new XElement(ns + "r", new XElement(ns + textName, "{na")),
                    new XElement(ns + "r", new XElement(ns + textName, "me}"))),
                new XElement(ns + containerName, new XElement(ns + textName, "{left")),
                new XElement(ns + containerName, new XElement(ns + textName, "right}")),
                new XElement(ns + containerName, new XElement(ns + textName, "{amount|if:A&B:Other}"))))
                .Save(xmlStream);
        }

        var tags = CliApplication.InspectTemplateTags(templatePath);

        Assert.Equal(new[] { "amount|if:A&B:Other", "name" }, tags);
    }

    [Fact]
    public void InspectTags_DoesNotJoinTextBoxTextWithItsOuterParagraph()
    {
        using var temp = new TemporaryDirectory();
        var templatePath = Path.Combine(temp.Path, "template.docx");
        var paragraph = new Paragraph(new Run(new Text("{outer")),
            new Run(new Picture(new DocumentFormat.OpenXml.Vml.Shape(
                new DocumentFormat.OpenXml.Vml.TextBox(new TextBoxContent(Paragraph("{inner}")))))),
            new Run(new Text("}")));
        File.WriteAllBytes(templatePath, CreateTemplate(paragraph));

        Assert.Equal(new[] { "inner", "outer" }, CliApplication.InspectTemplateTags(templatePath));
    }

    private static byte[] CreateTemplate(params OpenXmlElement[] bodyElements)
    {
        using (var stream = new MemoryStream())
        {
            using (var document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document, true))
            {
                var mainPart = document.AddMainDocumentPart();
                var body = new Body();

                foreach (var element in bodyElements)
                {
                    body.Append(element);
                }

                mainPart.Document = new Document(body);
                mainPart.Document.Save();
            }

            return stream.ToArray();
        }
    }

    private static Paragraph Paragraph(string text)
    {
        return new Paragraph(new Run(new Text(text)));
    }

    private static string ReadBodyText(byte[] docx)
    {
        using (var stream = new MemoryStream(docx))
        using (var document = WordprocessingDocument.Open(stream, false))
        {
            return string.Concat(document.MainDocumentPart!.Document.Body!.Descendants<Text>().Select(static text => text.Text));
        }
    }

    private sealed class TemporaryDirectory : IDisposable
    {
        public TemporaryDirectory()
        {
            Path = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "ndocx-templater-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(Path);
        }

        public string Path { get; }

        public void Dispose()
        {
            if (Directory.Exists(Path))
            {
                Directory.Delete(Path, recursive: true);
            }
        }
    }
}
