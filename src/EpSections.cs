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
    public sealed class EpSectionException : Exception
    {
        public EpSectionException(string message) : base(message) { }
    }

    public sealed class EpSectionInfo
    {
        public string Language { get; set; }
        public int DescriptionParagraphs { get; set; }
        public int ClaimParagraphs { get; set; }
        public string[] Notes { get; set; }
    }

    // The same section boundaries are used before Word compares the documents and
    // when measuring the saved comparison. Section counts therefore describe the
    // insertions actually highlighted in the single comparison document.
    public static class EpSections
    {
        private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        private static readonly XNamespace PackageR = "http://schemas.openxmlformats.org/package/2006/relationships";
        private static readonly XNamespace Types = "http://schemas.openxmlformats.org/package/2006/content-types";
        private const string DescriptionStyle = "LocalCompareEpDescription";
        private const string ClaimsStyle = "LocalCompareEpClaims";
        private const string DescriptionBookmark = "LocalCompareEpDescriptionBoundary";
        private const string ClaimsBookmark = "LocalCompareEpClaimsBoundary";

        public static EpSectionInfo Prepare(string sourceDocx, string outputDocx, string label)
        {
            string inputLabel = String.Equals(label, "Archive", StringComparison.OrdinalIgnoreCase) ? "Archive" :
                String.Equals(label, "New", StringComparison.OrdinalIgnoreCase) ? "New" : "EP";
            bool created = false;
            try
            {
                File.Copy(sourceDocx, outputDocx, false);
                created = true;
                File.SetAttributes(outputDocx, FileAttributes.Normal);
                using (var file = new FileStream(outputDocx, FileMode.Open, FileAccess.ReadWrite, FileShare.None))
                using (var zip = new ZipArchive(file, ZipArchiveMode.Update))
                {
                    XDocument document = Read(zip, "word/document.xml");
                    XElement body = document.Root == null ? null : document.Root.Element(W + "body");
                    if (body == null) throw Invalid(inputLabel, "does not contain a readable document body");
                    List<XElement> blocks = body.Elements().Where(e => e.Name != W + "sectPr").ToList();
                    List<int> descriptions = Indices(blocks, IsDescriptionHeading);
                    int description = descriptions.Count == 1 ? descriptions[0] : -1;
                    List<int> claims = description < 0 ? new List<int>() : Indices(blocks, IsClaimsHeading).Where(i => i > description).ToList();
                    if (Enumerable.Range(0, blocks.Count).Any(i => blocks[i].Descendants(W + "p").Any(p =>
                        (IsDescriptionHeading(p) || IsClaimsHeading(p) || IsExcludedHeading(p)) &&
                        !(description >= 0 && claims.Count > 0 && i > description && i < claims[0] &&
                          IsDescriptionColumnLabel(blocks[i], p)))))
                        throw Invalid(inputLabel, "contains a section heading inside a table, text box or content control, which cannot be separated reliably");
                    if (descriptions.Count != 1)
                        throw Invalid(inputLabel, "needs one standalone Description or Beschreibung heading");
                    if (claims.Count == 0)
                        throw Invalid(inputLabel, "needs a standalone Claims, Anspr\u00fcche, Patentanspr\u00fcche or Revendications heading after its description");
                    int claim = claims[0];
                    string language = ClaimLanguage(blocks[claim]);
                    if (claims.Skip(1).Any(i => ClaimLanguage(blocks[i]) == language))
                        throw Invalid(inputLabel, "contains repeated claims headings, so the original-language claims cannot be identified reliably");
                    int end = blocks.Count;
                    for (int i = claim + 1; i < blocks.Count; i++)
                    {
                        if (IsClaimsHeading(blocks[i]) || IsExcludedHeading(blocks[i])) { end = i; break; }
                    }
                    if (blocks.Skip(description + 1).Take(claim - description - 1).Any(IsExcludedHeading))
                        throw Invalid(inputLabel, "contains a drawings or references section before its claims");

                    List<XElement> descriptionBlocks = Content(blocks, description + 1, claim);
                    List<XElement> claimBlocks = Content(blocks, claim + 1, end);
                    if (!descriptionBlocks.Any(HasFinalWords) || !claimBlocks.Any(HasFinalWords))
                        throw Invalid(inputLabel, "needs nonempty description and original-language claims sections");
                    XElement firstClaim = claimBlocks.FirstOrDefault(HasFinalWords);
                    if (firstClaim == null || !LooksLikeFirstClaim(firstClaim))
                        throw Invalid(inputLabel, "does not have a recognizable first numbered claim after its claims heading");

                    XElement section = body.Elements(W + "sectPr").LastOrDefault();
                    XElement sectionCopy = section == null ? DefaultSection() : new XElement(section);
                    body.RemoveNodes();
                    body.Add(Marker("Description", DescriptionStyle, DescriptionBookmark, "2147483600"));
                    foreach (XElement block in descriptionBlocks) body.Add(block);
                    body.Add(Marker("Claims", ClaimsStyle, ClaimsBookmark, "2147483601"));
                    foreach (XElement block in claimBlocks) body.Add(block);
                    body.Add(sectionCopy);
                    StripRunningMatter(body);
                    EnsureStyles(zip);
                    FilterNotesAndRunningParts(zip, body);
                    Write(zip, "word/document.xml", document);
                    return new EpSectionInfo
                    {
                        Language = language,
                        DescriptionParagraphs = descriptionBlocks.SelectMany(Paragraphs).Count(p => HasFinalWords(p)),
                        ClaimParagraphs = claimBlocks.SelectMany(Paragraphs).Count(p => HasFinalWords(p)),
                        Notes = new string[0]
                    };
                }
            }
            catch (EpSectionException) { if (created) RemoveFailed(outputDocx); throw; }
            catch (Exception)
            {
                if (created) RemoveFailed(outputDocx);
                throw Invalid(inputLabel, "could not be read or prepared locally");
            }
        }

        // scope is Description, Claims or FullSpecification. Heading text is
        // removed from measurement copies, so full = description + claims.
        public static void ExtractScope(string sourceDocx, string outputDocx, string scope)
        {
            if (scope != "Description" && scope != "Claims" && scope != "FullSpecification")
                throw new ArgumentException("Unknown EP measurement section.", "scope");
            bool created = false;
            try
            {
                File.Copy(sourceDocx, outputDocx, false);
                created = true;
                File.SetAttributes(outputDocx, FileAttributes.Normal);
                using (var file = new FileStream(outputDocx, FileMode.Open, FileAccess.ReadWrite, FileShare.None))
                using (var zip = new ZipArchive(file, ZipArchiveMode.Update))
                {
                    XDocument document = Read(zip, "word/document.xml");
                    XElement body = document.Root == null ? null : document.Root.Element(W + "body");
                    if (body == null) throw new EpSectionException("The EP comparison does not contain a readable document body.");
                    List<XElement> blocks = body.Elements().Where(e => e.Name != W + "sectPr").ToList();
                    int description = FindMarker(blocks, DescriptionStyle, DescriptionBookmark, "Description");
                    int claim = FindMarker(blocks, ClaimsStyle, ClaimsBookmark, "Claims");
                    if (description < 0 || claim <= description)
                        throw new EpSectionException("Word did not preserve unambiguous EP section boundaries. No section figures can be reported.");
                    var keep = new List<XElement>();
                    if (scope == "Description" || scope == "FullSpecification")
                        keep.AddRange(blocks.Skip(description + 1).Take(claim - description - 1).Select(e => new XElement(e)));
                    if (scope == "Claims" || scope == "FullSpecification")
                        keep.AddRange(blocks.Skip(claim + 1).Select(e => new XElement(e)));
                    XElement section = body.Elements(W + "sectPr").LastOrDefault();
                    XElement sectionCopy = section == null ? DefaultSection() : new XElement(section);
                    body.RemoveNodes();
                    foreach (XElement block in keep) body.Add(block);
                    body.Add(sectionCopy);
                    StripRunningMatter(body);
                    FilterNotesAndRunningParts(zip, body);
                    Write(zip, "word/document.xml", document);
                }
            }
            catch (EpSectionException) { if (created) RemoveFailed(outputDocx); throw; }
            catch (Exception)
            {
                if (created) RemoveFailed(outputDocx);
                throw new EpSectionException("The EP comparison sections could not be measured locally.");
            }
        }

        private static EpSectionException Invalid(string label, string reason)
        {
            return new EpSectionException(label + " EP file " + reason + ". Prepare it with a standalone description heading followed by a standalone claims heading and the numbered original-language claims; remove cover pages, separate drawings and translated claims, then try again.");
        }

        private static List<int> Indices(List<XElement> blocks, Func<XElement, bool> predicate)
        {
            return Enumerable.Range(0, blocks.Count).Where(i => predicate(blocks[i])).ToList();
        }

        private static string Heading(XElement block)
        {
            if (block.Name != W + "p" || block.Descendants(W + "txbxContent").Any()) return String.Empty;
            return Regex.Replace(FinalText(block), @"\s+", " ").Trim().TrimEnd(':').Trim();
        }

        private static bool IsDescriptionHeading(XElement block)
        {
            string heading = Heading(block);
            return String.Equals(heading, "Description", StringComparison.OrdinalIgnoreCase) ||
                String.Equals(heading, "Beschreibung", StringComparison.OrdinalIgnoreCase);
        }

        private static bool IsClaimsHeading(XElement block) { return ClaimLanguage(block) != null; }

        // A non-leading Description column label in an ordinary data table is
        // content, not another section boundary. Only admit this narrow shape
        // between independently identified, standalone description/claims headings.
        // Ambiguous nested sections, claims/translations and exclusion headings
        // still stop preparation.
        private static bool IsDescriptionColumnLabel(XElement block, XElement paragraph)
        {
            if (block.Name != W + "tbl" || !IsDescriptionHeading(paragraph)) return false;
            XElement cell = paragraph.Parent;
            XElement row = cell == null ? null : cell.Parent;
            List<XElement> rows = block.Elements(W + "tr").ToList();
            if (cell == null || cell.Name != W + "tc" || row == null || row.Name != W + "tr" ||
                row.Parent != block || rows.Count < 2 || row != rows[0]) return false;
            List<XElement> cells = row.Elements(W + "tc").ToList();
            if (cells.Count < 2 || cells[0] == cell || cells.Count(c => c.Elements(W + "p").Any(IsDescriptionHeading)) != 1) return false;
            foreach (XElement headerCell in cells)
            {
                List<XElement> content = headerCell.Elements().Where(e => e.Name != W + "tcPr").ToList();
                if (content.Count != 1 || content[0].Name != W + "p" ||
                    content[0].Descendants(W + "p").Any() || content[0].Descendants(W + "drawing").Any() ||
                    content[0].Descendants(W + "pict").Any() || content[0].Descendants(W + "object").Any() ||
                    headerCell.Descendants(W + "gridSpan").Any() || headerCell.Descendants(W + "vMerge").Any() ||
                    headerCell.Descendants(W + "hMerge").Any()) return false;
                string label = Heading(content[0]);
                if (label.Length == 0 || label.Length > 80 || !HasFinalWords(content[0]) ||
                    IsClaimsHeading(content[0]) || IsExcludedHeading(content[0])) return false;
            }
            List<XElement> values = rows[1].Elements(W + "tc").ToList();
            return values.Count == cells.Count && values.All(HasFinalWords);
        }

        private static string ClaimLanguage(XElement block)
        {
            string heading = Heading(block);
            if (String.Equals(heading, "Claims", StringComparison.OrdinalIgnoreCase)) return "English";
            if (String.Equals(heading, "Anspr\u00fcche", StringComparison.OrdinalIgnoreCase) ||
                String.Equals(heading, "Patentanspr\u00fcche", StringComparison.OrdinalIgnoreCase)) return "German";
            if (String.Equals(heading, "Revendications", StringComparison.OrdinalIgnoreCase)) return "French";
            return null;
        }

        private static bool IsExcludedHeading(XElement block)
        {
            string heading = Heading(block);
            return new[] { "Drawing", "Drawings", "Figures", "Zeichnungen", "Dessins", "Cited references", "References cited", "References", "References cited in the description", "Zitierte Referenzen", "Documents cites", "Documents cit\u00e9s" }
                .Any(s => String.Equals(heading, s, StringComparison.OrdinalIgnoreCase));
        }

        private static List<XElement> Content(List<XElement> blocks, int first, int end)
        {
            var result = new List<XElement>();
            for (int i = first; i < end; i++)
            {
                XElement copy = new XElement(blocks[i]);
                foreach (XElement text in copy.Descendants(W + "t"))
                {
                    string normalized = text.Value.Replace('\u00a0', ' ').Replace('\u202f', ' ');
                    if (normalized != text.Value)
                    {
                        text.Value = normalized;
                        // Nonbreaking spaces can occupy a run of their own.
                        // Once normalized, Word drops leading/trailing ordinary
                        // spaces unless the text node explicitly preserves them.
                        if (normalized.StartsWith(" ", StringComparison.Ordinal) || normalized.EndsWith(" ", StringComparison.Ordinal))
                            text.SetAttributeValue(XNamespace.Xml + "space", "preserve");
                    }
                }
                // Empty paragraphs from PDF-to-Word conversion have no meaning
                // in a cropped specification. Keep images, tables and note refs.
                if (copy.Name == W + "p" && String.IsNullOrWhiteSpace(FinalText(copy)) &&
                    !copy.Descendants().Any(e => e.Name == W + "drawing" || e.Name == W + "pict" ||
                        e.Name == W + "object" || e.Name == W + "footnoteReference" || e.Name == W + "endnoteReference")) continue;
                result.Add(copy);
            }
            return result;
        }

        private static IEnumerable<XElement> Paragraphs(XElement block)
        {
            if (block.Name == W + "p") return new[] { block };
            return block.Descendants(W + "p");
        }

        private static bool HasFinalWords(XElement block) { return Regex.IsMatch(FinalText(block), @"[\p{L}\p{N}]"); }

        private static bool LooksLikeFirstClaim(XElement block)
        {
            XElement paragraph = block.Name == W + "p" ? block : block.Descendants(W + "p").FirstOrDefault(HasFinalWords);
            if (paragraph == null) return false;
            if (paragraph.Element(W + "pPr") != null && paragraph.Element(W + "pPr").Element(W + "numPr") != null) return true;
            return Regex.IsMatch(FinalText(paragraph), @"^\s*(?:1[\.)]|\(1\))\s*[\p{L}\p{N}]", RegexOptions.CultureInvariant);
        }

        private static string FinalText(XElement block)
        {
            var text = new StringBuilder();
            foreach (XElement node in block.Descendants())
            {
                if (node.Ancestors().Any(a => a.Name == W + "del" || a.Name == W + "moveFrom" ||
                    a.Name.LocalName.EndsWith("PrChange", StringComparison.Ordinal))) continue;
                if (node.Name == W + "t") text.Append(node.Value);
                else if (node.Name == W + "tab" || node.Name == W + "br" || node.Name == W + "cr") text.Append(' ');
            }
            return text.ToString();
        }

        private static XElement Marker(string text, string style, string bookmark, string id)
        {
            return new XElement(W + "p",
                new XElement(W + "pPr", new XElement(W + "pStyle", new XAttribute(W + "val", style))),
                new XElement(W + "bookmarkStart", new XAttribute(W + "id", id), new XAttribute(W + "name", bookmark)),
                new XElement(W + "r", new XElement(W + "t", text)),
                new XElement(W + "bookmarkEnd", new XAttribute(W + "id", id)));
        }

        private static int FindMarker(List<XElement> blocks, string style, string bookmark, string text)
        {
            // Formatting and bookmarks are useful independent signals, but Word
            // versions may discard either during native comparison. The retained
            // canonical heading is the required final-text check in every case.
            List<int> matching = Indices(blocks, p => Heading(p) == text);
            if (matching.Count != 1) return -1;
            int index = matching[0];
            return index;
        }

        private static void StripRunningMatter(XElement body)
        {
            foreach (XElement reference in body.Descendants().Where(e => e.Name == W + "headerReference" || e.Name == W + "footerReference").ToList()) reference.Remove();
        }

        private static XElement DefaultSection()
        {
            return new XElement(W + "sectPr", new XElement(W + "pgSz", new XAttribute(W + "w", "11906"), new XAttribute(W + "h", "16838")),
                new XElement(W + "pgMar", new XAttribute(W + "top", "1440"), new XAttribute(W + "right", "1440"), new XAttribute(W + "bottom", "1440"), new XAttribute(W + "left", "1440")));
        }

        private static void EnsureStyles(ZipArchive zip)
        {
            XDocument styles = zip.GetEntry("word/styles.xml") == null ?
                new XDocument(new XElement(W + "styles", new XAttribute(XNamespace.Xmlns + "w", W.NamespaceName))) : Read(zip, "word/styles.xml");
            foreach (string id in new[] { DescriptionStyle, ClaimsStyle })
            {
                foreach (XElement existing in styles.Root.Elements(W + "style").Where(e => (string)e.Attribute(W + "styleId") == id).ToList()) existing.Remove();
                styles.Root.Add(new XElement(W + "style", new XAttribute(W + "type", "paragraph"), new XAttribute(W + "customStyle", "1"), new XAttribute(W + "styleId", id),
                    new XElement(W + "name", new XAttribute(W + "val", id)),
                    new XElement(W + "pPr", new XElement(W + "keepNext"), new XElement(W + "spacing", new XAttribute(W + "before", "200"), new XAttribute(W + "after", "120"))),
                    new XElement(W + "rPr", new XElement(W + "b"))));
            }
            Write(zip, "word/styles.xml", styles);
            EnsurePart(zip, "/word/styles.xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml");
            const string relationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";
            XDocument relationships = zip.GetEntry("word/_rels/document.xml.rels") == null ? new XDocument(new XElement(PackageR + "Relationships")) : Read(zip, "word/_rels/document.xml.rels");
            if (!relationships.Root.Elements().Any(e => (string)e.Attribute("Type") == relationshipType))
            {
                string id = "rIdLocalCompareEpStyles";
                while (relationships.Root.Elements().Any(e => (string)e.Attribute("Id") == id)) id += "x";
                relationships.Root.Add(new XElement(PackageR + "Relationship", new XAttribute("Id", id), new XAttribute("Type", relationshipType), new XAttribute("Target", "styles.xml")));
                Write(zip, "word/_rels/document.xml.rels", relationships);
            }
        }

        private static void EnsurePart(ZipArchive zip, string part, string type)
        {
            XDocument contentTypes = Read(zip, "[Content_Types].xml");
            if (!contentTypes.Root.Elements(Types + "Override").Any(e => (string)e.Attribute("PartName") == part))
                contentTypes.Root.Add(new XElement(Types + "Override", new XAttribute("PartName", part), new XAttribute("ContentType", type)));
            Write(zip, "[Content_Types].xml", contentTypes);
        }

        private static void FilterNotesAndRunningParts(ZipArchive zip, XElement body)
        {
            foreach (string kind in new[] { "footnote", "endnote" })
            {
                string part = "word/" + kind + "s.xml";
                if (zip.GetEntry(part) == null) continue;
                var used = new HashSet<string>(body.Descendants(W + kind + "Reference")
                    .Where(e => !e.Ancestors().Any(a => a.Name == W + "del" || a.Name == W + "moveFrom"))
                    .Select(e => (string)e.Attribute(W + "id")), StringComparer.Ordinal);
                XDocument notes = Read(zip, part);
                foreach (XElement note in notes.Root.Elements(W + kind).ToList())
                {
                    int id;
                    string idText = (string)note.Attribute(W + "id");
                    if (Int32.TryParse(idText, out id) && id <= 0) continue;
                    if (!used.Contains(idText)) note.Remove();
                }
                Write(zip, part, notes);
            }
            List<string> removed = zip.Entries.Where(e => Regex.IsMatch(e.FullName, @"^word/(?:header|footer)\d+\.xml(?:\.rels)?$", RegexOptions.CultureInvariant) ||
                Regex.IsMatch(e.FullName, @"^word/_rels/(?:header|footer)\d+\.xml\.rels$", RegexOptions.CultureInvariant)).Select(e => e.FullName).ToList();
            foreach (string part in removed) zip.GetEntry(part).Delete();
            if (zip.GetEntry("[Content_Types].xml") != null)
            {
                XDocument types = Read(zip, "[Content_Types].xml");
                foreach (XElement part in types.Root.Elements(Types + "Override").Where(e => removed.Contains(((string)e.Attribute("PartName") ?? "").TrimStart('/'))).ToList()) part.Remove();
                Write(zip, "[Content_Types].xml", types);
            }
            if (zip.GetEntry("word/_rels/document.xml.rels") != null)
            {
                XDocument relationships = Read(zip, "word/_rels/document.xml.rels");
                foreach (XElement relationship in relationships.Root.Elements().Where(e =>
                    ((string)e.Attribute("Type") ?? "").EndsWith("/header", StringComparison.Ordinal) ||
                    ((string)e.Attribute("Type") ?? "").EndsWith("/footer", StringComparison.Ordinal)).ToList()) relationship.Remove();
                Write(zip, "word/_rels/document.xml.rels", relationships);
            }
        }

        private static XDocument Read(ZipArchive zip, string name)
        {
            ZipArchiveEntry entry = zip.GetEntry(name);
            if (entry == null) throw new InvalidDataException("Required document component missing.");
            using (Stream stream = entry.Open())
            using (XmlReader reader = XmlReader.Create(stream, new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null }))
                return XDocument.Load(reader, LoadOptions.PreserveWhitespace);
        }

        private static void Write(ZipArchive zip, string name, XDocument document)
        {
            ZipArchiveEntry existing = zip.GetEntry(name);
            if (existing != null) existing.Delete();
            using (Stream stream = zip.CreateEntry(name).Open())
            using (XmlWriter writer = XmlWriter.Create(stream, new XmlWriterSettings { Encoding = new UTF8Encoding(false) })) document.Save(writer);
        }

        private static void RemoveFailed(string path)
        {
            try { File.Delete(path); } catch { }
        }
    }
}
