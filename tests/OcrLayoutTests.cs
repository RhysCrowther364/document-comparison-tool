using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Web.Script.Serialization;
using System.Xml.Linq;
using LocalDocumentCompare;

public static class OcrLayoutTests
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private static string directory;
    private static int checks, nextFile;

    public static int Main()
    {
        directory = Path.Combine(Path.GetTempPath(), "LocalCompare-OcrLayoutTests-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        try
        {
            string text = "Heading & title <test>\r\nAlpha beta\u00a0gamma 123.\r\nIndented final line.\fLandscape page\nA second final paragraph.";
            string plain = Next("plain.docx");
            DocxContent.CreateFromRecognizedText(plain, text);
            object[] firstLines = { Box(300, 120, 1050, 55), Box(100, 220, 1250, 27), Box(210, 260, 1050, 27) };
            object[] secondLines = { Box(100, 100, 1300, 45), Box(100, 180, 1600, 27) };
            string layoutPath = WriteLayout(new { Version = 1, Pages = new[] {
                Page(816, 1056, 1700, 2200, firstLines), Page(1056, 816, 2200, 1700, secondLines) } });
            string formatted = Next("formatted.docx");
            Equal(true, OcrLayout.TryCreateDocument(formatted, text, layoutPath), "valid mixed-page layout applied");
            SameText(plain, formatted, "mixed-page formatting");
            XDocument xml = ReadDocument(formatted);
            XElement[] sections = xml.Descendants(W + "sectPr").ToArray();
            Equal(2, sections.Length, "one section per source page");
            Equal("12240", (string)sections[0].Element(W + "pgSz").Attribute(W + "w"), "portrait width from PDF DIPs");
            Equal("15840", (string)sections[0].Element(W + "pgSz").Attribute(W + "h"), "portrait height from PDF DIPs");
            Equal("15840", (string)sections[1].Element(W + "pgSz").Attribute(W + "w"), "landscape width retained");
            Equal("landscape", (string)sections[1].Element(W + "pgSz").Attribute(W + "orient"), "landscape orientation explicit");
            Equal("nextPage", (string)sections[0].Element(W + "type").Attribute(W + "val"), "section advances to next source page");
            Equal(0, xml.Descendants(W + "br").Count(e => (string)e.Attribute(W + "type") == "page"), "section break replaces page-break function");
            Equal(1, xml.Descendants(W + "br").Count(), "original whitespace-only separator retained");
            XElement[] textParagraphs = xml.Descendants(W + "p").Where(p => p.Descendants(W + "t").Any()).ToArray();
            Equal(5, textParagraphs.Length, "one paragraph per recognised line");
            Equal(true, textParagraphs.All(p => (string)p.Element(W + "pPr").Element(W + "spacing").Attribute(W + "after") == "0"), "explicit zero paragraph spacing after");
            Equal(true, textParagraphs.All(p => (string)p.Element(W + "pPr").Element(W + "spacing").Attribute(W + "lineRule") == "atLeast"), "text line spacing avoids forced glyph clipping");
            Equal(true, textParagraphs.SelectMany(p => p.Elements(W + "r")).All(r => (string)r.Element(W + "rPr").Element(W + "rFonts").Attribute(W + "ascii") == "Arial"), "consistent readable approximate font");
            Equal(true, FontSize(textParagraphs[0]) > FontSize(textParagraphs[1]), "larger heading retained as larger text");
            Equal(true, Int32.Parse((string)textParagraphs[2].Element(W + "pPr").Element(W + "ind").Attribute(W + "left")) > 0, "source indentation retained");
            Equal(true, Int32.Parse((string)textParagraphs[1].Element(W + "pPr").Element(W + "spacing").Attribute(W + "before")) > 0, "visible vertical gap approximated");
            Equal(0, xml.Descendants().Count(e => e.Name.LocalName == "txbxContent" || e.Name.LocalName == "drawing" || e.Name.LocalName == "tbl"), "no duplicate image text or inferred floating structures");

            string columns = "Left column top\nLeft column bottom\nRight column top\nRight column bottom";
            string columnsPlain = Next("columns-plain.docx");
            DocxContent.CreateFromRecognizedText(columnsPlain, columns);
            string columnsLayout = WriteLayout(new { Version = 1, Pages = new[] { Page(816, 1056, 1700, 2200,
                new[] { Box(100, 100, 600, 27), Box(100, 145, 600, 27), Box(900, 100, 600, 27), Box(900, 145, 600, 27) }) } });
            string columnsOutput = Next("columns.docx");
            Equal(true, OcrLayout.TryCreateDocument(columnsOutput, columns, columnsLayout), "column geometry accepted without reordering");
            SameText(columnsPlain, columnsOutput, "columns retain established OCR order");
            XElement[] columnParagraphs = ReadDocument(columnsOutput).Descendants(W + "p").ToArray();
            Equal("0", (string)columnParagraphs[2].Element(W + "pPr").Element(W + "spacing").Attribute(W + "before"), "column reset does not invent negative or giant spacing");
            Equal(true, Int32.Parse((string)columnParagraphs[2].Element(W + "pPr").Element(W + "ind").Attribute(W + "left")) >
                Int32.Parse((string)columnParagraphs[0].Element(W + "pPr").Element(W + "ind").Attribute(W + "left")), "right-column text retains its horizontal cue");

            string sparseLayout = WriteLayout(new { Version = 1, Pages = new[] { Page(816, 1056, 1700, 2200,
                new[] { Box(400, 800, 900, 27) }) } });
            string sparseOutput = Next("sparse.docx");
            Equal(true, OcrLayout.TryCreateDocument(sparseOutput, "Sparse central caption.", sparseLayout), "sparse page geometry applied");
            XElement sparseParagraph = ReadDocument(sparseOutput).Descendants(W + "p").First();
            Equal("4320", (string)sparseParagraph.Element(W + "pPr").Element(W + "spacing").Attribute(W + "before"), "first line retains vertical position beyond capped top margin");
            Equal("Sparse central caption.\n", DocxContent.ReadFinalText(sparseOutput), "sparse-page placement does not change text");

            string oneLine = "Alpha beta gamma.";
            string validOne = Serialize(new { Version = 1, Pages = new[] { Page(816, 1056, 1700, 2200, new[] { Box(100, 100, 600, 27) }) } });
            Fallback(oneLine, null, "missing sidecar");
            Fallback(oneLine, "{broken", "malformed JSON");
            Fallback(oneLine, "null", "null metadata");
            Fallback(oneLine, validOne.Replace("\"Version\":1", "\"Version\":2"), "unsupported version");
            Fallback(oneLine, "{\"Version\":1,\"Pages\":[]}", "page-count mismatch");
            Fallback(oneLine, Serialize(new { Version = 1, Pages = new[] { Page(816, 1056, 1700, 2200, new object[0]) } }), "line-count mismatch");
            Fallback(oneLine, validOne.Replace("\"X\":100", "\"X\":-1"), "negative coordinate");
            Fallback(oneLine, validOne.Replace("\"X\":100", "\"X\":1699"), "box extends outside image");
            Fallback(oneLine, validOne.Replace("\"ImageWidth\":1700", "\"ImageWidth\":0"), "zero image dimension");
            Fallback(oneLine, validOne.Replace("\"Width\":816", "\"Width\":3000"), "paper larger than Word limit");
            Fallback(oneLine, validOne.Replace("\"Height\":27", "\"Height\":0"), "empty word box");
            Fallback(oneLine, validOne.Replace("\"Y\":100", "\"Y\":null"), "missing coordinate");
            Fallback(oneLine, validOne.Replace("\"Height\":27", "\"Height\":\"NaN\""), "non-finite geometry");
            Fallback("", "{\"Version\":1,\"Pages\":[]}", "empty OCR text still preserved");
            Fallback("Alpha.\f", Serialize(new { Version = 1, Pages = new[] {
                Page(816, 1056, 1700, 2200, new[] { Box(100, 100, 600, 27) }), Page(816, 1056, 1700, 2200, new object[0]) } }), "blank source page falls back without dropping its boundary");

            string existing = Next("existing.docx");
            DocxContent.CreateFromRecognizedText(existing, "Original retained text.");
            string existingBefore = DocxContent.ReadFinalText(existing);
            bool refused = false;
            try { OcrLayout.TryCreateDocument(existing, oneLine, layoutPath); }
            catch (IOException) { refused = true; }
            Equal(true, refused, "existing destination is not overwritten");
            Equal(existingBefore, DocxContent.ReadFinalText(existing), "existing output remains intact");
            Console.WriteLine("OcrLayoutTests passed: " + checks + " checks.");
            return 0;
        }
        catch (Exception error) { Console.Error.WriteLine(error); return 1; }
        finally { Directory.Delete(directory, true); }
    }

    private static object Box(double x, double y, double width, double height)
    { return new { X = x, Y = y, Width = width, Height = height }; }
    private static object Page(double width, double height, double imageWidth, double imageHeight, object[] lines)
    { return new { Width = width, Height = height, ImageWidth = imageWidth, ImageHeight = imageHeight, Lines = lines }; }
    private static string Serialize(object value) { return new JavaScriptSerializer().Serialize(value); }
    private static string WriteLayout(object value)
    {
        string path = Next("layout.json");
        File.WriteAllText(path, Serialize(value), new UTF8Encoding(false));
        return path;
    }
    private static void Fallback(string text, string metadata, string scenario)
    {
        string sidecar = Next("fallback.json");
        if (metadata != null) File.WriteAllText(sidecar, metadata, new UTF8Encoding(false));
        string plain = Next("fallback-plain.docx"), output = Next("fallback.docx");
        DocxContent.CreateFromRecognizedText(plain, text);
        Equal(false, OcrLayout.TryCreateDocument(output, text, sidecar), scenario + " uses plain fallback");
        SameText(plain, output, scenario);
        Equal(ReadDocument(plain).ToString(), ReadDocument(output).ToString(), scenario + " leaves plain formatting intact");
    }
    private static void SameText(string expected, string actual, string scenario)
    {
        TextMetrics baseline = DocxContent.Analyze(expected), result = DocxContent.Analyze(actual);
        Equal(baseline.Words, result.Words, scenario + " preserves word count");
        Equal(baseline.InsertedWords, result.InsertedWords, scenario + " preserves insertion count");
        Equal(baseline.TextFingerprint, result.TextFingerprint, scenario + " preserves strict text fingerprint");
        Equal(DocxContent.ReadFinalText(expected), DocxContent.ReadFinalText(actual), scenario + " preserves exact visible text");
    }
    private static int FontSize(XElement paragraph)
    { return Int32.Parse((string)paragraph.Element(W + "r").Element(W + "rPr").Element(W + "sz").Attribute(W + "val")); }
    private static XDocument ReadDocument(string path)
    {
        using (var zip = ZipFile.OpenRead(path))
        using (var stream = zip.GetEntry("word/document.xml").Open()) return XDocument.Load(stream, LoadOptions.PreserveWhitespace);
    }
    private static string Next(string suffix) { return Path.Combine(directory, (++nextFile).ToString("D3") + "-" + suffix); }
    private static void Equal<T>(T expected, T actual, string label)
    {
        checks++;
        if (!Object.Equals(expected, actual)) throw new Exception(label + ": expected " + expected + ", got " + actual + ".");
    }
}
