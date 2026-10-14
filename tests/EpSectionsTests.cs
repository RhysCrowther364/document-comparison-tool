using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using LocalDocumentCompare;

public static class EpSectionsTests
{
    private const string Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private static readonly XNamespace W = Namespace;
    private static int checks;

    public static int Main(string[] args)
    {
        string directory = Path.Combine(Path.GetTempPath(), "LocalCompare-EpTest-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        try
        {
            if (args.Length != 0 && args.Length != 2)
                throw new ArgumentException("Optional sample checks require an archive path and a prepared path.");
            string raw = Path.Combine(directory, "raw.docx");
            string normalized = Path.Combine(directory, "normalized.docx");
            NewDocx(raw, P("Confidential cover material") + P("Description") + P("Alpha\u00a0beta") + P("") +
                P("Claims") + P("1. Alpha gamma.") + P("Anspr\u00fcche") + P("1. German translation.") +
                P("Revendications") + P("1. French translation.") + P("Drawing") + P("Drawing text") + P("Cited references") + P("References text"));
            string before = Hash(raw);
            EpSectionInfo info = EpSections.Prepare(raw, normalized, "Archive");
            Equal("English", info.Language, "original first claims language");
            Equal(1, info.DescriptionParagraphs, "description paragraph count");
            Equal(1, info.ClaimParagraphs, "claims paragraph count");
            Equal(before, Hash(raw), "source remains unchanged");
            TextMetrics description = Measure(normalized, directory, "Description");
            TextMetrics claims = Measure(normalized, directory, "Claims");
            TextMetrics full = Measure(normalized, directory, "FullSpecification");
            Equal(2, description.Words, "description counts exclude headings and blank paragraphs");
            Equal(3, claims.Words, "claims count includes original numbered claim only");
            Equal(5, full.Words, "full specification excludes cover, translations, drawings and references");
            Equal(description.Words + claims.Words, full.Words, "full scope is description plus claims");
            Equal(0, full.InsertedWords, "prepared sources have no inserted text");
            using (var zip = ZipFile.OpenRead(normalized))
            {
                Equal(2, Read(zip, "word/document.xml").Descendants(W + "bookmarkStart").Count(), "boundary bookmarks written");
                Equal(2, Read(zip, "word/styles.xml").Descendants(W + "style").Count(), "boundary paragraph styles written");
            }

            string german = Path.Combine(directory, "german.docx");
            NewDocx(german, P("Beschreibung:") + P("Ein Text.") + P("Patentanspr\u00fcche:") + P("1) Ein Anspruch.") + P("Claims") + P("1. A translation."));
            Equal("German", EpSections.Prepare(german, Path.Combine(directory, "german-prepared.docx"), "New").Language, "German original claims");
            string french = Path.Combine(directory, "french.docx");
            NewDocx(french, P("DESCRIPTION") + P("Un texte.") + P("REVENDICATIONS") + P("(1) Une revendication.") + P("Drawings") + P("discarded illustration text"));
            Equal("French", EpSections.Prepare(french, Path.Combine(directory, "french-prepared.docx"), "New").Language, "French original claims");
            string automatic = Path.Combine(directory, "automatic.docx");
            NewDocx(automatic, P("Description") + P("Alpha beta") + P("Claims") + "<w:p><w:pPr><w:numPr><w:ilvl w:val='0'/><w:numId w:val='1'/></w:numPr></w:pPr><w:r><w:t>Automatically numbered claim.</w:t></w:r></w:p>");
            Equal("English", EpSections.Prepare(automatic, Path.Combine(directory, "automatic-prepared.docx"), "New").Language, "automatic Word numbering accepted");

            string dataTable = "<w:tbl><w:tr><w:tc>" + P("Abbreviation") + "</w:tc><w:tc>" + P("Component Type") +
                "</w:tc><w:tc>" + P("Description") + "</w:tc></w:tr><w:tr><w:tc>" + P("MM") +
                "</w:tc><w:tc>" + P("Type A") + "</w:tc><w:tc>" + P("Alpha resin") +
                "</w:tc></w:tr><w:tr><w:tc>" + P("XX") + "</w:tc><w:tc>" + P("Type B") +
                "</w:tc><w:tc>" + P("Beta resin") + "</w:tc></w:tr></w:tbl>";
            string tableInput = Path.Combine(directory, "description-column.docx");
            string tableOutput = Path.Combine(directory, "description-column-prepared.docx");
            NewDocx(tableInput, P("Description") + P("Alpha beta") + dataTable + P("Claims") + P("1. A claim.") +
                P("Anspr\u00fcche") + P("1. Translated claim."));
            string tableBefore = Hash(tableInput);
            Equal("English", EpSections.Prepare(tableInput, tableOutput, "New").Language, "Description data-table label accepted within description");
            Equal(16, Measure(tableOutput, directory, "Description").Words, "data-table headings and all rows retained in description count");
            Equal(3, Measure(tableOutput, directory, "Claims").Words, "data-table label does not change original claims selection");
            Equal(19, Measure(tableOutput, directory, "FullSpecification").Words, "full specification retains entire data table and excludes translation");
            Equal(tableBefore, Hash(tableInput), "data-table source remains unchanged");
            using (var zip = ZipFile.OpenRead(tableOutput))
                Equal(XElement.Parse("<root xmlns:w='" + Namespace + "'>" + dataTable + "</root>").Elements().Single().ToString(SaveOptions.DisableFormatting),
                    Read(zip, "word/document.xml").Descendants(W + "tbl").Single().ToString(SaveOptions.DisableFormatting), "data table XML remains unchanged");
            Reject(directory, "column-before-description", dataTable + P("Description") + P("Alpha beta") + P("Claims") + P("1. Claim"), "heading inside");
            Reject(directory, "column-in-claims", P("Description") + P("Alpha beta") + P("Claims") + P("1. Claim") + dataTable, "heading inside");
            Reject(directory, "column-missing-description", P("Alpha beta") + dataTable + P("Claims") + P("1. Claim"), "heading inside");
            Reject(directory, "column-missing-claims", P("Description") + P("Alpha beta") + dataTable, "heading inside");
            Reject(directory, "column-first-cell", P("Description") + P("Alpha beta") + dataTable.Replace(P("Abbreviation"), P("Description")).Replace(P("Component Type"), P("Type")).Replace(P("Description") + "</w:tc></w:tr>", P("Details") + "</w:tc></w:tr>") + P("Claims") + P("1. Claim"), "heading inside");
            Reject(directory, "column-claims-label", P("Description") + P("Alpha beta") + dataTable.Replace(P("Component Type"), P("Claims")) + P("Claims") + P("1. Claim"), "heading inside");
            Reject(directory, "column-multiple-paragraphs", P("Description") + P("Alpha beta") + dataTable.Replace(P("Description"), P("Description") + P("Another section")) + P("Claims") + P("1. Claim"), "heading inside");
            Reject(directory, "column-merged", P("Description") + P("Alpha beta") + dataTable.Replace("<w:tc>" + P("Description"), "<w:tc><w:tcPr><w:gridSpan w:val='2'/></w:tcPr>" + P("Description")) + P("Claims") + P("1. Claim"), "heading inside");
            Reject(directory, "column-content-control", P("Description") + P("Alpha beta") + "<w:sdt><w:sdtContent>" + dataTable + "</w:sdtContent></w:sdt>" + P("Claims") + P("1. Claim"), "heading inside");

            string spacedInput = Path.Combine(directory, "nonbreaking-runs.docx");
            string spacedOutput = Path.Combine(directory, "nonbreaking-runs-prepared.docx");
            NewDocx(spacedInput, P("Description") + "<w:p><w:r><w:t>Alpha</w:t></w:r><w:r><w:t>&#xA0;</w:t></w:r>" +
                "<w:r><w:t>beta</w:t></w:r><w:r><w:t>&#x202F;gamma&#xA0;</w:t></w:r><w:r><w:t>delta</w:t></w:r></w:p>" +
                P("Claims") + P("1. A claim."));
            string spacedBefore = Hash(spacedInput);
            EpSections.Prepare(spacedInput, spacedOutput, "New");
            using (var zip = ZipFile.OpenRead(spacedOutput))
            {
                XDocument xml = Read(zip, "word/document.xml");
                foreach (string value in new[] { " ", " gamma " })
                    Equal("preserve", (string)xml.Descendants(W + "t").Single(t => t.Value == value).Attribute(XNamespace.Xml + "space"), "normalized run-edge whitespace remains preserved in Word");
            }
            Equal(4, Measure(spacedOutput, directory, "Description").Words, "nonbreaking run boundaries keep all four words distinct");
            Equal(7, Measure(spacedOutput, directory, "FullSpecification").Words, "normalized spaces do not join words across runs");
            Equal(spacedBefore, Hash(spacedInput), "normalizing working-copy spacing leaves original unchanged");

            Reject(directory, "missing", P("Description") + P("Alpha beta"), "needs a standalone Claims");
            Reject(directory, "repeated", P("Description") + P("Alpha beta") + P("Claims") + P("1. Claim") + P("Claims") + P("1. Duplicate"), "repeated claims headings");
            Reject(directory, "empty", P("Description") + P(" ") + P("Claims") + P("1. Claim"), "nonempty");
            Reject(directory, "unnumbered", P("Description") + P("Alpha beta") + P("Claims") + P("Unnumbered ambiguous prose"), "first numbered claim");
            Reject(directory, "table-heading", "<w:tbl><w:tr><w:tc>" + P("Description") + P("Alpha beta") + "</w:tc></w:tr></w:tbl>" + P("Claims") + P("1. Claim"), "heading inside");
            Reject(directory, "nested-original-claims", P("Description") + P("Alpha beta") + "<w:tbl><w:tr><w:tc>" + P("Anspr\u00fcche") + P("1. Original claim") + "</w:tc></w:tr></w:tbl>" + P("Claims") + P("1. Translation"), "heading inside");
            Reject(directory, "nested-references", P("Description") + P("Alpha beta") + P("Claims") + P("1. Original claim") + "<w:tbl><w:tr><w:tc>" + P("References") + P("Unrelated bibliography") + "</w:tc></w:tr></w:tbl>", "heading inside");
            Reject(directory, "textbox-heading", P("Description") + P("Alpha beta") + "<w:p><w:r><w:pict><w:txbxContent>" + P("Anspr\u00fcche") + P("1. Original claim") + "</w:txbxContent></w:pict></w:r></w:p>" + P("Claims") + P("1. Translation"), "heading inside");
            Reject(directory, "sdt-heading", P("Description") + P("Alpha beta") + "<w:sdt><w:sdtContent>" + P("Claims") + P("1. Original claim") + "</w:sdtContent></w:sdt>", "heading inside");
            Reject(directory, "wrong-order", P("Description") + P("Alpha beta") + P("Drawing") + P("Claims") + P("1. Claim"), "before its claims");

            string compared = Path.Combine(directory, "compared.docx");
            File.Copy(normalized, compared);
            using (var zip = ZipFile.Open(compared, ZipArchiveMode.Update))
            {
                XDocument xml = Read(zip, "word/document.xml");
                XElement body = xml.Root.Element(W + "body");
                XElement desc = body.Elements(W + "p").First(p => String.Concat(p.Descendants(W + "t").Select(t => t.Value)) == "Alpha beta");
                desc.Add(new XElement(W + "ins", new XAttribute(W + "id", "1"), new XElement(W + "r", new XElement(W + "t", new XAttribute(XNamespace.Xml + "space", "preserve"), " inserted"))));
                desc.Add(new XElement(W + "del", new XAttribute(W + "id", "2"), new XElement(W + "r", new XElement(W + "delText", " deleted archive text"))));
                XElement claim = body.Elements(W + "p").First(p => String.Concat(p.Descendants(W + "t").Select(t => t.Value)) == "1. Alpha gamma.");
                claim.Add(new XElement(W + "ins", new XAttribute(W + "id", "3"), new XElement(W + "r", new XElement(W + "t", new XAttribute(XNamespace.Xml + "space", "preserve"), " added words"))));
                claim.Add(new XElement(W + "moveTo", new XAttribute(W + "id", "4"), new XElement(W + "r", new XElement(W + "t", new XAttribute(XNamespace.Xml + "space", "preserve"), " retained move"))));
                Write(zip, "word/document.xml", xml);
            }
            description = Measure(compared, directory, "Description");
            claims = Measure(compared, directory, "Claims");
            full = Measure(compared, directory, "FullSpecification");
            Equal(1, description.InsertedWords, "description metrics retain actual insertion markup");
            Equal(2, claims.InsertedWords, "claims metrics exclude moved text from inserted words");
            Equal(3, full.InsertedWords, "full highlighted words equal section highlighted words");
            Equal(description.Words + claims.Words, full.Words, "comparison final-text counts retain deletions and moves correctly");

            string running = Path.Combine(directory, "running.docx");
            NewDocx(running, P("Description") + "<w:p><w:r><w:t>Alpha beta</w:t><w:footnoteReference w:id='1'/></w:r></w:p>" + P("Claims") + P("1. Claim text"));
            AddRunningNotes(running);
            string runningOutput = Path.Combine(directory, "running-prepared.docx");
            EpSections.Prepare(running, runningOutput, "New");
            Equal(4, Measure(runningOutput, directory, "Description").Words, "referenced description footnote retained");
            Equal(3, Measure(runningOutput, directory, "Claims").Words, "description footnote omitted from claims count");
            Equal(7, Measure(runningOutput, directory, "FullSpecification").Words, "unused notes and running headers do not count");
            using (var zip = ZipFile.OpenRead(runningOutput))
            {
                Equal(null, zip.GetEntry("word/header1.xml"), "header part removed");
                Equal(0, Read(zip, "word/document.xml").Descendants(W + "headerReference").Count(), "header reference removed");
                Equal(0, Read(zip, "word/_rels/document.xml.rels").Root.Elements().Count(e => ((string)e.Attribute("Type") ?? "").EndsWith("/header")), "header relationship removed");
                Equal(0, Read(zip, "[Content_Types].xml").Root.Elements().Count(e => (string)e.Attribute("PartName") == "/word/header1.xml"), "header content type removed");
            }

            // Word can omit bookmarks or remap styles: canonical retained headings
            // still establish boundaries; duplicated final headings must fail.
            string noMarkers = Path.Combine(directory, "no-markers.docx");
            File.Copy(compared, noMarkers);
            using (var zip = ZipFile.Open(noMarkers, ZipArchiveMode.Update))
            {
                XDocument xml = Read(zip, "word/document.xml");
                foreach (XElement marker in xml.Descendants().Where(e => e.Name == W + "bookmarkStart" || e.Name == W + "bookmarkEnd" || e.Name == W + "pStyle").ToList()) marker.Remove();
                Write(zip, "word/document.xml", xml);
            }
            Equal(3, Measure(noMarkers, directory, "FullSpecification").InsertedWords, "canonical headings survive omitted marker metadata");
            string ambiguous = Path.Combine(directory, "ambiguous.docx");
            File.Copy(compared, ambiguous);
            using (var zip = ZipFile.Open(ambiguous, ZipArchiveMode.Update))
            {
                XDocument xml = Read(zip, "word/document.xml");
                xml.Root.Element(W + "body").AddFirst(XElement.Parse(P("Claims")));
                Write(zip, "word/document.xml", xml);
            }
            string failedMeasurement = Path.Combine(directory, "ambiguous-measurement.docx");
            try { EpSections.ExtractScope(ambiguous, failedMeasurement, "Description"); throw new Exception("Ambiguous comparison boundary was accepted."); }
            catch (EpSectionException) { checks++; }
            Equal(false, File.Exists(failedMeasurement), "failed measurement copy removed");

            string existing = Path.Combine(directory, "existing.docx");
            File.WriteAllText(existing, "existing user content");
            string existingHash = Hash(existing);
            try { EpSections.Prepare(raw, existing, "New"); throw new Exception("Existing output was overwritten."); }
            catch (EpSectionException) { checks++; }
            Equal(existingHash, Hash(existing), "pre-existing output is preserved");

            string sampleRaw = args.Length == 2 ? args[0] : null;
            string samplePrep = args.Length == 2 ? args[1] : null;
            if (args.Length == 2)
            {
                if (!File.Exists(sampleRaw) || !File.Exists(samplePrep))
                    throw new ArgumentException("Both optional sample files must exist.");
                string rawHash = Hash(sampleRaw), prepHash = Hash(samplePrep);
                string normalizedRaw = Path.Combine(directory, "sample-raw.docx"), normalizedPrep = Path.Combine(directory, "sample-prep.docx");
                EpSectionInfo rawInfo = EpSections.Prepare(sampleRaw, normalizedRaw, "Archive");
                EpSectionInfo prepInfo = EpSections.Prepare(samplePrep, normalizedPrep, "New");
                Equal(true, rawInfo.DescriptionParagraphs > 0, "sample raw description has content");
                Equal(true, rawInfo.ClaimParagraphs > 0, "sample raw original claims have content");
                Equal(rawInfo.DescriptionParagraphs, prepInfo.DescriptionParagraphs, "sample prepared description count");
                Equal(rawInfo.ClaimParagraphs, prepInfo.ClaimParagraphs, "sample prepared claims count");
                foreach (string scope in new[] { "Description", "Claims", "FullSpecification" })
                {
                    TextMetrics rawMetrics = Measure(normalizedRaw, directory, scope);
                    TextMetrics prepMetrics = Measure(normalizedPrep, directory, scope);
                    Equal(rawMetrics.Words, prepMetrics.Words, "sample " + scope + " raw/prepared word count");
                    Equal(rawMetrics.TextFingerprint, prepMetrics.TextFingerprint, "sample " + scope + " raw/prepared text equivalence");
                }
                Equal(rawHash, Hash(sampleRaw), "user raw sample unchanged");
                Equal(prepHash, Hash(samplePrep), "user prepped sample unchanged");
                Console.WriteLine("Sample raw/prepared section equivalence passed.");
            }
            else Console.WriteLine("No optional sample paths supplied; portable synthetic checks completed.");
            Console.WriteLine("PASS: " + checks + " EP boundary, scope, revision, language and source-preservation checks.");
            return 0;
        }
        catch (Exception ex) { Console.Error.WriteLine(ex.Message); return 1; }
        finally
        {
            string cleanup = Path.GetFullPath(directory);
            string temporaryRoot = Path.GetFullPath(Path.GetTempPath()).TrimEnd(Path.DirectorySeparatorChar) + Path.DirectorySeparatorChar;
            if (!cleanup.StartsWith(temporaryRoot, StringComparison.OrdinalIgnoreCase) ||
                !Path.GetFileName(cleanup).StartsWith("LocalCompare-EpTest-", StringComparison.Ordinal))
                throw new InvalidOperationException("Unexpected EP test cleanup path.");
            Directory.Delete(cleanup, true);
        }
    }

    private static TextMetrics Measure(string source, string directory, string scope)
    {
        string path = Path.Combine(directory, "measurement-" + Guid.NewGuid().ToString("N") + ".docx");
        EpSections.ExtractScope(source, path, scope);
        return DocxContent.Analyze(path);
    }

    private static void Reject(string directory, string name, string body, string reason)
    {
        string input = Path.Combine(directory, name + ".docx");
        string output = Path.Combine(directory, name + "-prepared.docx");
        NewDocx(input, body);
        try { EpSections.Prepare(input, output, "New"); throw new Exception(name + " malformed EP input was accepted."); }
        catch (EpSectionException ex)
        {
            if (!ex.Message.Contains(reason)) throw new Exception(name + " did not provide a useful preparation reason.");
            checks++;
        }
        Equal(false, File.Exists(output), name + " failed normalization copy removed");
    }

    private static void AddRunningNotes(string path)
    {
        using (var zip = ZipFile.Open(path, ZipArchiveMode.Update))
        {
            XDocument document = Read(zip, "word/document.xml");
            document.Root.Element(W + "body").Element(W + "sectPr").Add(new XElement(W + "headerReference", new XAttribute(W + "type", "default"), new XAttribute(XName.Get("id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships"), "rIdHeader")));
            Write(zip, "word/document.xml", document);
            Write(zip, "word/header1.xml", XDocument.Parse("<w:hdr xmlns:w='" + Namespace + "'>" + P("Cover header should never be counted") + "</w:hdr>"));
            Write(zip, "word/footnotes.xml", XDocument.Parse("<w:footnotes xmlns:w='" + Namespace + "'><w:footnote w:id='0'>" + P("separator") + "</w:footnote><w:footnote w:id='1'>" + P("body note") + "</w:footnote><w:footnote w:id='2'>" + P("discarded cover note with unrelated words") + "</w:footnote></w:footnotes>"));
            XNamespace relations = "http://schemas.openxmlformats.org/package/2006/relationships";
            Write(zip, "word/_rels/document.xml.rels", new XDocument(new XElement(relations + "Relationships", new XElement(relations + "Relationship", new XAttribute("Id", "rIdHeader"), new XAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header"), new XAttribute("Target", "header1.xml")), new XElement(relations + "Relationship", new XAttribute("Id", "rIdFootnote"), new XAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes"), new XAttribute("Target", "footnotes.xml")))));
            XDocument types = Read(zip, "[Content_Types].xml");
            XNamespace ct = "http://schemas.openxmlformats.org/package/2006/content-types";
            types.Root.Add(new XElement(ct + "Override", new XAttribute("PartName", "/word/header1.xml"), new XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml")), new XElement(ct + "Override", new XAttribute("PartName", "/word/footnotes.xml"), new XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml")));
            Write(zip, "[Content_Types].xml", types);
        }
    }

    private static void NewDocx(string path, string body)
    {
        DocxContent.CreateFromRecognizedText(path, String.Empty);
        using (var zip = ZipFile.Open(path, ZipArchiveMode.Update))
            Write(zip, "word/document.xml", XDocument.Parse("<w:document xmlns:w='" + Namespace + "'><w:body>" + body + "<w:sectPr/></w:body></w:document>"));
    }

    private static string P(string text)
    {
        return new XElement(W + "p", new XElement(W + "r", new XElement(W + "t", new XAttribute(XNamespace.Xml + "space", "preserve"), text))).ToString(SaveOptions.DisableFormatting);
    }

    private static string Hash(string path)
    {
        using (var sha = SHA256.Create()) using (var stream = File.OpenRead(path)) return Convert.ToBase64String(sha.ComputeHash(stream));
    }

    private static XDocument Read(ZipArchive zip, string name)
    {
        using (Stream stream = zip.GetEntry(name).Open())
        using (XmlReader reader = XmlReader.Create(stream, new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null })) return XDocument.Load(reader);
    }

    private static void Write(ZipArchive zip, string name, XDocument xml)
    {
        var existing = zip.GetEntry(name); if (existing != null) existing.Delete();
        using (var writer = XmlWriter.Create(zip.CreateEntry(name).Open(), new XmlWriterSettings { Encoding = new UTF8Encoding(false) })) xml.Save(writer);
    }

    private static void Equal<T>(T expected, T actual, string label)
    {
        if (!Object.Equals(expected, actual)) throw new Exception(label + ": expected " + expected + ", got " + actual);
        checks++;
    }
}
