using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using LocalDocumentCompare;

public static class EngineTests
{
    private const string Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private static int checks;
    public static int Main()
    {
        string directory = Path.Combine(Path.GetTempPath(), "LocalCompare-MetricsTest-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        try
        {
            string source = Path.Combine(directory, "sample.docx");
            string body =
                "<w:p><w:r><w:t>for</w:t></w:r><w:r><w:rPr><w:b/></w:rPr><w:t>mat</w:t></w:r>" +
                "<w:r><w:t xml:space='preserve'> keep </w:t></w:r><w:ins w:id='1'><w:r><w:t>new</w:t></w:r></w:ins>" +
                "<w:del w:id='2'><w:r><w:delText> removed</w:delText></w:r></w:del></w:p>" +
                "<w:p><w:moveFrom w:id='3'><w:r><w:t>moved words</w:t></w:r></w:moveFrom></w:p>" +
                "<w:p><w:moveTo w:id='4'><w:r><w:t>moved words</w:t></w:r></w:moveTo></w:p>" +
                "<w:p><w:r><w:t>chang</w:t></w:r><w:ins w:id='5'><w:r><w:t>ed</w:t></w:r></w:ins></w:p>" +
                "<w:tbl><w:tr><w:trPr><w:ins w:id='6'/></w:trPr><w:tc><w:p><w:r><w:t>table addition</w:t></w:r></w:p></w:tc></w:tr>" +
                "<w:tr><w:trPr><w:del w:id='7'/></w:trPr><w:tc><w:p><w:r><w:t>deleted row</w:t></w:r></w:p></w:tc></w:tr></w:tbl>" +
                "<w:p><w:r><w:instrText>DO NOT COUNT FIELD CODE</w:instrText></w:r><w:r><w:t>field</w:t></w:r></w:p>" +
                "<w:p><w:r><w:t>can't co-operate 2026</w:t></w:r></w:p>";
            using (var zip = ZipFile.Open(source, ZipArchiveMode.Create))
            {
                Add(zip, "word/document.xml", "<w:document xmlns:w='" + Namespace + "'><w:body>" + body + "</w:body></w:document>");
                Add(zip, "word/header1.xml", "<w:hdr xmlns:w='" + Namespace + "'><w:p><w:r><w:t>header</w:t></w:r></w:p></w:hdr>");
                Add(zip, "word/footnotes.xml", "<w:footnotes xmlns:w='" + Namespace + "'><w:footnote w:id='0'><w:p><w:r><w:t>separator</w:t></w:r></w:p></w:footnote><w:footnote w:id='1'><w:p><w:r><w:t>note</w:t></w:r></w:p></w:footnote></w:footnotes>");
                Add(zip, "word/comments.xml", "<w:comments xmlns:w='" + Namespace + "'><w:p><w:r><w:t>ignore comment</w:t></w:r></w:p></w:comments>");
                Add(zip, "word/settings.xml", "<w:settings xmlns:w='" + Namespace + "'><w:trackRevisions/></w:settings>");
            }
            TextMetrics metric = DocxContent.Analyze(source);
            Equal(14, metric.Words, "final word count and scopes");
            Equal(4, metric.InsertedWords, "partial word and inserted row; moved words retained");
            string languageText = DocxContent.ReadFinalText(source);
            Equal(true, languageText.Contains("format keep new"), "language text joins formatted runs and keeps final insertions");
            Equal(false, languageText.Contains("removed") || languageText.Contains("deleted row"), "language text excludes tracked deletions");
            Equal(false, languageText.Contains("DO NOT COUNT FIELD CODE") || languageText.Contains("header") || languageText.Contains("note"),
                "language text excludes field instructions and separate running/note parts");
            DocxContent.HighlightInsertionsAndShowMarkup(source);
            TextMetrics highlighted = DocxContent.Analyze(source);
            Equal(metric.TextFingerprint, highlighted.TextFingerprint, "highlight preserves final text");
            using (var zip = ZipFile.OpenRead(source))
            {
                XNamespace w = Namespace;
                XDocument document = Read(zip, "word/document.xml");
                Equal(3, document.Descendants(w + "highlight").Count(), "insertions highlighted across word/row revisions");
                Equal(1, document.Descendants(w + "moveFrom").Count(), "move origin preserved");
                Equal(1, document.Descendants(w + "moveTo").Count(), "move destination preserved");
                Equal(2, document.Descendants(w + "del").Count(), "deletions preserved");
                Equal(0, Read(zip, "word/settings.xml").Descendants(w + "trackRevisions").Count(), "result starts with tracking disabled");
            }
            string external = Path.Combine(directory, "external.docx");
            using (var zip = ZipFile.Open(external, ZipArchiveMode.Create))
            {
                Add(zip, "word/document.xml", "<w:document xmlns:w='" + Namespace + "' xmlns:r='http://schemas.openxmlformats.org/officeDocument/2006/relationships'><w:body><w:p><w:hyperlink r:id='rId1'><w:r><w:t>cached label</w:t></w:r></w:hyperlink><w:r><w:fldChar w:fldCharType='begin'/></w:r></w:p></w:body></w:document>");
                Add(zip, "word/_rels/document.xml.rels", "<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'><Relationship Id='rId1' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink' Target='https://example.invalid/' TargetMode='External'/></Relationships>");
            }
            Equal(true, DocxContent.DetachExternalRelationships(external), "external link detected");
            using (var zip = ZipFile.OpenRead(external))
            {
                XNamespace w = Namespace;
                XNamespace r = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
                XDocument document = Read(zip, "word/document.xml");
                Equal(0, document.Descendants().Attributes(r + "id").Count(), "external ID removed from owner");
                Equal(0, Read(zip, "word/_rels/document.xml.rels").Root.Elements().Count(), "external relationship removed");
                Equal("1", (string)document.Descendants(w + "fldChar").First().Attribute(w + "fldLock"), "cached field locked");
            }
            Equal(2, DocxContent.Analyze(external).Words, "cached hyperlink text preserved");
            string joined = Sample(directory, "joined", "<w:p><w:r><w:t>alpha beta gamma</w:t></w:r></w:p>");
            const string deletedMark = "<w:pPr><w:rPr><w:del w:id='40'/></w:rPr></w:pPr>";
            string trackedJoin = Sample(directory, "tracked-join",
                "<w:p>" + deletedMark + "<w:r><w:t xml:space='preserve'>alpha </w:t></w:r></w:p>" +
                "<w:p>" + deletedMark + "<w:r><w:t xml:space='preserve'>beta </w:t></w:r></w:p>" +
                "<w:p><w:r><w:t>gamma</w:t></w:r></w:p>");
            TextMetrics trackedJoinMetrics = DocxContent.Analyze(trackedJoin);
            Equal(3, trackedJoinMetrics.Words, "deleted paragraph marks keep final words");
            Equal(DocxContent.Analyze(joined).TextFingerprint, trackedJoinMetrics.TextFingerprint,
                "deleted paragraph marks produce the joined final paragraph fingerprint");
            string word = Sample(directory, "whole-word", "<w:p><w:r><w:t>architecture</w:t></w:r></w:p>");
            string wordJoin = Sample(directory, "word-join",
                "<w:p>" + deletedMark + "<w:r><w:t>archi</w:t></w:r></w:p>" +
                "<w:p><w:ins w:id='41'><w:r><w:t>tecture</w:t></w:r></w:ins></w:p>");
            TextMetrics wordJoinMetrics = DocxContent.Analyze(wordJoin);
            Equal(1, wordJoinMetrics.Words, "deleting a paragraph mark can join word pieces");
            Equal(1, wordJoinMetrics.InsertedWords, "joined word retains insertion flags");
            Equal(DocxContent.Analyze(word).TextFingerprint, wordJoinMetrics.TextFingerprint, "joined word fingerprint");
            DocxContent.HighlightInsertionsAndShowMarkup(wordJoin);
            Equal(wordJoinMetrics.TextFingerprint, DocxContent.Analyze(wordJoin).TextFingerprint, "highlighting preserves joined final text");
            using (var zip = ZipFile.OpenRead(wordJoin))
            {
                XNamespace w = Namespace;
                XDocument document = Read(zip, "word/document.xml");
                Equal(1, document.Descendants(w + "del").Count(), "paragraph-mark deletion revision remains in output");
                Equal(1, document.Descendants(w + "ins").Count(), "joined insertion revision remains in output");
            }
            string mismatch = Sample(directory, "mismatch", "<w:p><w:r><w:t>alpha beta delta</w:t></w:r></w:p>");
            Equal(false, DocxContent.Analyze(mismatch).TextFingerprint == trackedJoinMetrics.TextFingerprint,
                "genuine changed token still fails reconciliation");
            string reordered = Sample(directory, "reordered", "<w:p><w:r><w:t>beta alpha gamma</w:t></w:r></w:p>");
            Equal(false, DocxContent.Analyze(reordered).TextFingerprint == trackedJoinMetrics.TextFingerprint,
                "token order within a final paragraph remains checked");
            string separate = Sample(directory, "ordinary-paragraphs",
                "<w:p><w:r><w:t>archi</w:t></w:r></w:p><w:p><w:r><w:t>tecture</w:t></w:r></w:p>");
            Equal(2, DocxContent.Analyze(separate).Words, "ordinary paragraph marks keep words separate");
            Equal(false, DocxContent.Analyze(separate).TextFingerprint == wordJoinMetrics.TextFingerprint,
                "ordinary paragraphs retain distinct fingerprint streams");
            string cells = Sample(directory, "cell-boundaries", "<w:tbl><w:tr><w:tc><w:p>" + deletedMark +
                "<w:r><w:t>archi</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>tecture</w:t></w:r></w:p></w:tc></w:tr></w:tbl>");
            Equal(2, DocxContent.Analyze(cells).Words, "deleted paragraph mark cannot join words across cells");
            Equal(DocxContent.Analyze(separate).TextFingerprint, DocxContent.Analyze(cells).TextFingerprint,
                "cell boundaries retain distinct fingerprint streams");
            string formerMark = Sample(directory, "old-property-revision",
                "<w:p><w:pPr><w:pPrChange w:id='42'><w:pPr><w:rPr><w:del w:id='43'/></w:rPr></w:pPr></w:pPrChange></w:pPr>" +
                "<w:r><w:t>archi</w:t></w:r></w:p><w:p><w:r><w:t>tecture</w:t></w:r></w:p>");
            Equal(2, DocxContent.Analyze(formerMark).Words, "previous paragraph properties do not delete the current mark");
            const string movedMark = "<w:pPr><w:rPr><w:moveFrom w:id='50'/></w:rPr></w:pPr>";
            string movedJoin = Sample(directory, "moved-paragraph-marks",
                "<w:p>" + movedMark + "<w:r><w:t xml:space='preserve'>alpha </w:t></w:r></w:p>" +
                "<w:p>" + movedMark + "<w:r><w:t xml:space='preserve'>beta </w:t></w:r></w:p>" +
                "<w:p><w:r><w:t>gamma</w:t></w:r></w:p>");
            TextMetrics movedJoinMetrics = DocxContent.Analyze(movedJoin);
            Equal(3, movedJoinMetrics.Words, "moved-away paragraph marks keep final words");
            Equal(0, movedJoinMetrics.InsertedWords, "moving paragraph marks does not insert surviving words");
            Equal(DocxContent.Analyze(joined).TextFingerprint, movedJoinMetrics.TextFingerprint,
                "moved-away paragraph marks produce the joined final paragraph fingerprint");
            Equal("alpha beta gamma\n", DocxContent.ReadFinalText(movedJoin),
                "language extraction joins moved-away paragraph marks");
            string mixedJoin = Sample(directory, "mixed-deleted-and-moved-marks",
                "<w:p>" + deletedMark + "<w:r><w:t xml:space='preserve'>alpha </w:t></w:r></w:p>" +
                "<w:p>" + movedMark + "<w:r><w:t xml:space='preserve'>beta </w:t></w:r></w:p>" +
                "<w:p>" + deletedMark + "<w:ins w:id='51'><w:r><w:t xml:space='preserve'>gamma </w:t></w:r></w:ins></w:p>" +
                "<w:p><w:r><w:t>delta</w:t></w:r></w:p>");
            string mixedFinal = Sample(directory, "mixed-final", "<w:p><w:r><w:t>alpha beta gamma delta</w:t></w:r></w:p>");
            TextMetrics mixedMetrics = DocxContent.Analyze(mixedJoin);
            Equal(4, mixedMetrics.Words, "mixed consecutive deleted and moved marks retain all final words");
            Equal(1, mixedMetrics.InsertedWords, "mixed paragraph joins retain insertion flags");
            Equal(DocxContent.Analyze(mixedFinal).TextFingerprint, mixedMetrics.TextFingerprint,
                "mixed consecutive paragraph marks reconcile with the final joined paragraph");
            string movedWordJoin = Sample(directory, "moved-word-join",
                "<w:p>" + movedMark + "<w:r><w:t>archi</w:t></w:r></w:p>" +
                "<w:p><w:ins w:id='52'><w:r><w:t>tecture</w:t></w:r></w:ins></w:p>");
            TextMetrics movedWordMetrics = DocxContent.Analyze(movedWordJoin);
            Equal(1, movedWordMetrics.Words, "moved-away paragraph mark can join word pieces");
            Equal(1, movedWordMetrics.InsertedWords, "partial word insertion survives a moved paragraph mark");
            Equal(DocxContent.Analyze(word).TextFingerprint, movedWordMetrics.TextFingerprint,
                "moved-away word-piece join reconciles with the complete final word");
            Equal("architecture\n", DocxContent.ReadFinalText(movedWordJoin),
                "language extraction joins word pieces across moved-away marks");
            DocxContent.HighlightInsertionsAndShowMarkup(movedWordJoin);
            Equal(movedWordMetrics.TextFingerprint, DocxContent.Analyze(movedWordJoin).TextFingerprint,
                "highlighting preserves final text joined across moved-away marks");
            using (var zip = ZipFile.OpenRead(movedWordJoin))
            {
                XNamespace w = Namespace;
                XDocument document = Read(zip, "word/document.xml");
                Equal(1, document.Descendants(w + "moveFrom").Count(), "moved-away paragraph revision remains in output");
                Equal(1, document.Descendants(w + "highlight").Count(), "joined word insertion receives yellow highlighting");
            }
            string movedToMark = Sample(directory, "surviving-move-destination-mark",
                "<w:p><w:pPr><w:rPr><w:moveTo w:id='53'/></w:rPr></w:pPr><w:r><w:t>archi</w:t></w:r></w:p>" +
                "<w:p><w:r><w:t>tecture</w:t></w:r></w:p>");
            Equal(2, DocxContent.Analyze(movedToMark).Words, "move destination paragraph mark survives and separates word pieces");
            Equal(DocxContent.Analyze(separate).TextFingerprint, DocxContent.Analyze(movedToMark).TextFingerprint,
                "move destination mark keeps the separate final paragraph streams");
            string formerMovedMark = Sample(directory, "historical-moved-mark",
                "<w:p><w:pPr><w:pPrChange w:id='54'><w:pPr><w:rPr><w:moveFrom w:id='55'/></w:rPr></w:pPr></w:pPrChange></w:pPr>" +
                "<w:r><w:t>archi</w:t></w:r></w:p><w:p><w:r><w:t>tecture</w:t></w:r></w:p>");
            Equal(2, DocxContent.Analyze(formerMovedMark).Words, "historical paragraph property snapshot does not move the current mark");
            Equal(DocxContent.Analyze(separate).TextFingerprint, DocxContent.Analyze(formerMovedMark).TextFingerprint,
                "historical moved paragraph mark does not change final paragraph boundaries");
            string formerRunMovedMark = Sample(directory, "historical-mark-run-properties",
                "<w:p><w:pPr><w:rPr><w:rPrChange w:id='56'><w:rPr><w:moveFrom w:id='57'/></w:rPr></w:rPrChange></w:rPr></w:pPr>" +
                "<w:r><w:t>archi</w:t></w:r></w:p><w:p><w:r><w:t>tecture</w:t></w:r></w:p>");
            Equal(2, DocxContent.Analyze(formerRunMovedMark).Words, "historical paragraph-mark run properties do not move the current mark");
            Equal(DocxContent.Analyze(separate).TextFingerprint, DocxContent.Analyze(formerRunMovedMark).TextFingerprint,
                "historical moved mark inside a run-property snapshot does not join paragraphs");
            string movedCells = Sample(directory, "moved-mark-cell-boundaries", "<w:tbl><w:tr><w:tc><w:p>" + movedMark +
                "<w:r><w:t>archi</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>tecture</w:t></w:r></w:p></w:tc></w:tr></w:tbl>");
            Equal(2, DocxContent.Analyze(movedCells).Words, "moved-away paragraph mark cannot join word pieces across cells");
            Equal(DocxContent.Analyze(separate).TextFingerprint, DocxContent.Analyze(movedCells).TextFingerprint,
                "cell boundaries remain separate after moved-away paragraph marks");
            string movedTextBox = Sample(directory, "moved-mark-textbox-boundaries",
                "<w:p>" + movedMark + "<w:r><w:t>archi</w:t></w:r></w:p>" +
                "<w:p><w:r><w:pict><w:txbxContent><w:p>" + movedMark +
                "<w:r><w:t>tecture</w:t></w:r></w:p></w:txbxContent></w:pict></w:r></w:p>");
            Equal(2, DocxContent.Analyze(movedTextBox).Words, "moved-away paragraph mark cannot join body and text-box words");
            Equal(DocxContent.Analyze(separate).TextFingerprint, DocxContent.Analyze(movedTextBox).TextFingerprint,
                "text-box boundaries retain distinct final text streams after moved-away marks");
            Console.WriteLine("PASS: " + checks + " metric, revision, highlight and external-link checks.");
            return 0;
        }
        catch (Exception ex) { Console.Error.WriteLine(ex.Message); return 1; }
        finally { Directory.Delete(directory, true); }
    }
    private static void Equal<T>(T expected, T actual, string name)
    {
        if (!Object.Equals(expected, actual)) throw new Exception(name + ": expected " + expected + ", got " + actual);
        checks++;
    }
    private static void Add(ZipArchive zip, string name, string content)
    {
        using (var writer = new StreamWriter(zip.CreateEntry(name).Open(), new UTF8Encoding(false))) writer.Write(content);
    }
    private static string Sample(string directory, string name, string body)
    {
        string path = Path.Combine(directory, name + ".docx");
        using (var zip = ZipFile.Open(path, ZipArchiveMode.Create))
            Add(zip, "word/document.xml", "<w:document xmlns:w='" + Namespace + "'><w:body>" + body + "</w:body></w:document>");
        return path;
    }
    private static XDocument Read(ZipArchive zip, string name)
    {
        using (var stream = zip.GetEntry(name).Open()) return XDocument.Load(stream);
    }
}
