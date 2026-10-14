using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;
using System.Xml.Linq;

namespace LocalDocumentCompare
{
    // Reads the final text represented by Word's comparison, without accepting its revisions.
    public sealed class TextMetrics
    {
        public int Words { get; set; }
        public int InsertedWords { get; set; }
        public bool HasImages { get; set; }
        public string TextFingerprint { get; set; }
    }

    public static class DocxContent
    {
        private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        private static readonly XNamespace R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        private static readonly XNamespace PackageR = "http://schemas.openxmlformats.org/package/2006/relationships";
        private static readonly Regex Word = new Regex(@"[\p{L}\p{N}][\p{L}\p{M}\p{N}]*(?:['’\-][\p{L}\p{N}][\p{L}\p{M}\p{N}]*)*", RegexOptions.Compiled);

        public static void CreateFromRecognizedText(string path, string text)
        {
            XNamespace types = "http://schemas.openxmlformats.org/package/2006/content-types";
            var body = new XElement(W + "body");
            string[] pages = text.Replace("\r\n", "\n").Replace('\r', '\n').Split('\f');
            for (int page = 0; page < pages.Length; page++)
            {
                if (page > 0) body.Add(new XElement(W + "p", new XElement(W + "r", new XElement(W + "br", new XAttribute(W + "type", "page")))));
                foreach (string line in pages[page].Split('\n'))
                    body.Add(new XElement(W + "p", new XElement(W + "r", new XElement(W + "t", new XAttribute(XNamespace.Xml + "space", "preserve"), line))));
            }
            body.Add(new XElement(W + "sectPr",
                new XElement(W + "pgSz", new XAttribute(W + "w", "11906"), new XAttribute(W + "h", "16838")),
                new XElement(W + "pgMar", new XAttribute(W + "top", "1440"), new XAttribute(W + "right", "1440"), new XAttribute(W + "bottom", "1440"), new XAttribute(W + "left", "1440"))));
            using (var zip = ZipFile.Open(path, ZipArchiveMode.Create))
            {
                WriteNew(zip, "[Content_Types].xml", new XDocument(new XElement(types + "Types",
                    new XElement(types + "Default", new XAttribute("Extension", "rels"), new XAttribute("ContentType", "application/vnd.openxmlformats-package.relationships+xml")),
                    new XElement(types + "Default", new XAttribute("Extension", "xml"), new XAttribute("ContentType", "application/xml")),
                    new XElement(types + "Override", new XAttribute("PartName", "/word/document.xml"), new XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml")))));
                WriteNew(zip, "_rels/.rels", new XDocument(new XElement(PackageR + "Relationships",
                    new XElement(PackageR + "Relationship", new XAttribute("Id", "rId1"), new XAttribute("Type", R.NamespaceName + "/officeDocument"), new XAttribute("Target", "word/document.xml")))));
                WriteNew(zip, "word/document.xml", new XDocument(new XElement(W + "document", new XAttribute(XNamespace.Xmlns + "w", W.NamespaceName), body)));
            }
        }

        private static void WriteNew(ZipArchive zip, string name, XDocument xml)
        {
            using (Stream stream = zip.CreateEntry(name).Open())
            using (XmlWriter writer = XmlWriter.Create(stream, new XmlWriterSettings { Encoding = new UTF8Encoding(false) })) xml.Save(writer);
        }

        public static TextMetrics Analyze(string path)
        {
            var result = new TextMetrics();
            // Word can reorder parts, so normalize the collection of per-paragraph token streams.
            var streams = new List<string>();
            // Word retains a write handle to its saved working copy while it is open.
            // Read the completed save without denying Word that existing handle.
            using (var file = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var zip = new ZipArchive(file, ZipArchiveMode.Read))
            {
                foreach (var entry in zip.Entries.Where(e => IsTextPart(e.FullName)).OrderBy(e => e.FullName, StringComparer.Ordinal))
                {
                    XDocument xml = Load(entry);
                    result.HasImages |= xml.Descendants().Any(e => e.Name == W + "drawing" || e.Name == W + "pict" || e.Name.LocalName == "blip");
                    var text = new StringBuilder();
                    var inserted = new List<bool>();
                    Action finish = delegate
                    {
                        if (text.Length == 0) return;
                        var words = new List<string>();
                        foreach (Match match in Word.Matches(text.ToString()))
                        {
                            result.Words++;
                            bool isInserted = false;
                            for (int i = match.Index; i < match.Index + match.Length; i++) isInserted |= inserted[i];
                            if (isInserted) result.InsertedWords++;
                            words.Add(match.Value);
                        }
                        if (words.Count > 0) streams.Add(String.Join("\u001F", words));
                        text.Length = 0;
                        inserted.Clear();
                    };
                    Walk(xml.Root, false, text, inserted, finish);
                    finish();
                }
            }
            streams.Sort(StringComparer.Ordinal);
            using (var hash = System.Security.Cryptography.SHA256.Create())
                result.TextFingerprint = Convert.ToBase64String(hash.ComputeHash(Encoding.UTF8.GetBytes(String.Join("\u001E", streams))));
            return result;
        }

        // Language detection uses only final visible main-document text, with
        // the same revision and paragraph handling as word measurement. Headers,
        // footers, comments, field instructions and proofing metadata are not
        // source-language evidence. The text stays in memory and is never logged.
        public static string ReadFinalText(string path)
        {
            using (var file = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var zip = new ZipArchive(file, ZipArchiveMode.Read))
            {
                ZipArchiveEntry entry = zip.GetEntry("word/document.xml");
                if (entry == null) throw new InvalidDataException("Required document component missing.");
                XDocument xml = Load(entry);
                var result = new StringBuilder();
                var paragraph = new StringBuilder();
                var marks = new List<bool>();
                Action finish = delegate
                {
                    if (paragraph.Length > 0) result.Append(paragraph).Append('\n');
                    paragraph.Length = 0;
                    marks.Clear();
                };
                Walk(xml.Root, false, paragraph, marks, finish);
                finish();
                return result.ToString();
            }
        }

        private static bool IsTextPart(string name)
        {
            return name == "word/document.xml" || name == "word/footnotes.xml" || name == "word/endnotes.xml" ||
                Regex.IsMatch(name, @"^word/(header|footer)\d+\.xml$", RegexOptions.CultureInvariant);
        }

        private static bool Deleted(XElement e)
        {
            if (e.Name == W + "del" || e.Name == W + "moveFrom") return true;
            if (e.Name == W + "tr")
            {
                var properties = e.Element(W + "trPr");
                if (properties != null && properties.Element(W + "del") != null) return true;
            }
            if (e.Name == W + "tc")
            {
                var properties = e.Element(W + "tcPr");
                if (properties != null && properties.Element(W + "cellDel") != null) return true;
            }
            if (e.Name == W + "footnote" || e.Name == W + "endnote")
            {
                int id;
                if (Int32.TryParse((string)e.Attribute(W + "id"), out id) && id <= 0) return true;
            }
            return false;
        }

        private static bool Inserted(XElement e, bool inherited)
        {
            if (e.Name == W + "ins") return true;
            if (e.Name == W + "moveTo") return false;
            if (e.Name == W + "tr")
            {
                var p = e.Element(W + "trPr");
                if (p != null && p.Element(W + "ins") != null) return true;
            }
            if (e.Name == W + "tc")
            {
                var p = e.Element(W + "tcPr");
                if (p != null && p.Element(W + "cellIns") != null) return true;
            }
            return inherited;
        }

        private static void Walk(XElement element, bool isInserted, StringBuilder text, List<bool> marks, Action finish)
        {
            if (element == null || Deleted(element)) return;
            // Skip metadata/property snapshots, comments and non-displayed field instructions.
            if (element.Name.Namespace == W && (element.Name.LocalName.EndsWith("Pr", StringComparison.Ordinal) ||
                element.Name.LocalName.EndsWith("PrChange", StringComparison.Ordinal) || element.Name.LocalName == "instrText" ||
                element.Name.LocalName == "delInstrText" || element.Name.LocalName == "delText")) return;
            if (element.Name.LocalName == "AlternateContent")
            {
                // Choice and Fallback represent the same content. Do not count both.
                XElement chosen = element.Elements().FirstOrDefault(e => e.Name.LocalName == "Choice") ?? element.Elements().FirstOrDefault();
                if (chosen != null) Walk(chosen, isInserted, text, marks, finish);
                return;
            }
            isInserted = Inserted(element, isInserted);
            bool paragraph = element.Name == W + "p";
            bool boundary = element.Name == W + "tc" || element.Name == W + "txbxContent" ||
                element.Name == W + "footnote" || element.Name == W + "endnote";
            if (boundary) finish();
            if (element.Name == W + "t")
            {
                string value = element.Value;
                text.Append(value);
                for (int i = 0; i < value.Length; i++) marks.Add(isInserted);
            }
            else if (element.Name == W + "tab" || element.Name == W + "br" || element.Name == W + "cr")
            {
                text.Append(' '); marks.Add(isInserted);
            }
            else if (element.Name == W + "noBreakHyphen")
            {
                text.Append('-'); marks.Add(isInserted);
            }
            else
            {
                foreach (XElement child in element.Elements()) Walk(child, isInserted, text, marks, finish);
            }
            // A deleted or moved-away paragraph mark merges this paragraph's
            // surviving text with the following paragraph. Keep its characters
            // and insertion flags together until a surviving mark is reached;
            // adding an artificial separator here can also split a single word.
            // Cell, text-box and note boundaries still terminate their own text.
            if (boundary || paragraph && !HasDeletedParagraphMark(element)) finish();
        }

        private static bool HasDeletedParagraphMark(XElement paragraph)
        {
            XElement properties = paragraph.Element(W + "pPr");
            XElement markProperties = properties == null ? null : properties.Element(W + "rPr");
            return markProperties != null && (markProperties.Element(W + "del") != null ||
                markProperties.Element(W + "moveFrom") != null);
        }

        public static void HighlightInsertionsAndShowMarkup(string path)
        {
            using (var zip = ZipFile.Open(path, ZipArchiveMode.Update))
            {
                foreach (var entry in zip.Entries.Where(e => IsTextPart(e.FullName)).ToArray())
                {
                    XDocument xml = Load(entry);
                    Highlight(xml.Root, false);
                    Replace(entry, xml);
                }
                var settings = zip.GetEntry("word/settings.xml");
                if (settings != null)
                {
                    XDocument xml = Load(settings);
                    XElement view = xml.Root.Element(W + "revisionView");
                    if (view == null) { view = new XElement(W + "revisionView"); xml.Root.Add(view); }
                    foreach (string attribute in new[] { "markup", "insDel", "formatting", "inkAnnotations" })
                        view.SetAttributeValue(W + attribute, "1");
                    // Opening in Word should show the stored comparison without automatically tracking later edits.
                    xml.Root.Elements(W + "trackRevisions").Remove();
                    Replace(settings, xml);
                }
            }
        }

        private static void Highlight(XElement element, bool inserted)
        {
            if (element == null || Deleted(element)) return;
            if (element.Name.Namespace == W && (element.Name.LocalName.EndsWith("Pr", StringComparison.Ordinal) ||
                element.Name.LocalName.EndsWith("PrChange", StringComparison.Ordinal))) return;
            inserted = Inserted(element, inserted);
            if (element.Name == W + "r" && inserted && element.Descendants(W + "t").Any())
            {
                XElement properties = element.Element(W + "rPr");
                if (properties == null) { properties = new XElement(W + "rPr"); element.AddFirst(properties); }
                properties.Elements(W + "highlight").Remove();
                properties.Add(new XElement(W + "highlight", new XAttribute(W + "val", "yellow")));
            }
            foreach (XElement child in element.Elements().ToArray()) Highlight(child, inserted);
        }

        // External linked objects/templates must not be fetched while opening a temporary OOXML copy.
        // Cached display text/images remain available. This does not inspect legacy binary .doc files.
        public static bool DetachExternalRelationships(string path)
        {
            bool changed = false;
            using (var zip = ZipFile.Open(path, ZipArchiveMode.Update))
            {
                foreach (ZipArchiveEntry entry in zip.Entries.Where(e => e.FullName.EndsWith(".rels", StringComparison.OrdinalIgnoreCase)).ToArray())
                {
                    XDocument rels = Load(entry);
                    var external = rels.Root.Elements(PackageR + "Relationship").Where(e => String.Equals((string)e.Attribute("TargetMode"), "External", StringComparison.OrdinalIgnoreCase)).ToArray();
                    if (external.Length == 0) continue;
                    changed = true;
                    var ids = new HashSet<string>(external.Select(e => (string)e.Attribute("Id")), StringComparer.Ordinal);
                    string name = entry.FullName;
                    int relFolder = name.LastIndexOf("/_rels/", StringComparison.Ordinal);
                    string ownerName = relFolder < 0 ? null : name.Substring(0, relFolder + 1) + name.Substring(relFolder + 7, name.Length - relFolder - 12);
                    ZipArchiveEntry owner = ownerName == null ? null : zip.GetEntry(ownerName);
                    if (owner != null)
                    {
                        XDocument xml = Load(owner);
                        foreach (XElement element in xml.Descendants().ToArray())
                        {
                            var attributes = element.Attributes().Where(a => a.Name.Namespace == R && ids.Contains(a.Value)).ToArray();
                            if (attributes.Length == 0) continue;
                            if (element.Name == W + "attachedTemplate" || element.Name == W + "altChunk" || element.Name.LocalName == "OLEObject")
                                element.Remove();
                            else
                                foreach (var attribute in attributes) attribute.Remove();
                        }
                        Replace(owner, xml);
                    }
                    foreach (XElement externalRelationship in external) externalRelationship.Remove();
                    Replace(entry, rels);
                }
                // DDE/INCLUDETEXT and similar instructions could fetch files. Keep their cached results only.
                foreach (ZipArchiveEntry entry in zip.Entries.Where(e => IsTextPart(e.FullName)).ToArray())
                {
                    XDocument xml = Load(entry);
                    bool fieldsChanged = false;
                    foreach (XElement field in xml.Descendants(W + "fldChar"))
                    {
                        field.SetAttributeValue(W + "fldLock", "1");
                        field.SetAttributeValue(W + "dirty", "0");
                        fieldsChanged = true;
                    }
                    foreach (XElement field in xml.Descendants(W + "fldSimple"))
                    {
                        field.SetAttributeValue(W + "fldLock", "1");
                        field.SetAttributeValue(W + "dirty", "0");
                        fieldsChanged = true;
                    }
                    if (fieldsChanged) Replace(entry, xml);
                }
            }
            return changed;
        }

        private static XDocument Load(ZipArchiveEntry entry)
        {
            using (Stream stream = entry.Open())
            using (XmlReader reader = XmlReader.Create(stream, new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null }))
                return XDocument.Load(reader, LoadOptions.PreserveWhitespace);
        }

        private static void Replace(ZipArchiveEntry entry, XDocument xml)
        {
            using (Stream stream = entry.Open())
            {
                stream.SetLength(0);
                using (XmlWriter writer = XmlWriter.Create(stream, new XmlWriterSettings { Encoding = new UTF8Encoding(false), CloseOutput = false, Indent = false }))
                    xml.Save(writer);
            }
        }
    }
}
