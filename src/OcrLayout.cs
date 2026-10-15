using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Web.Script.Serialization;
using System.Xml;
using System.Xml.Linq;

namespace LocalDocumentCompare
{
    // Geometry is optional. Recognised text, line order and page boundaries are
    // always taken from the existing OCR text file, never reconstructed from boxes.
    public static class OcrLayout
    {
        private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        private const int MaximumMetadataBytes = 32 * 1024 * 1024;

        public static bool TryCreateDocument(string outputPath, string canonicalText, string layoutPath)
        {
            DocxContent.CreateFromRecognizedText(outputPath, canonicalText);
            string[][] lines = canonicalText.Replace("\r\n", "\n").Replace('\r', '\n').Split('\f')
                .Select(page => page.Split('\n')).ToArray();
            LayoutData layout;
            if (!TryReadLayout(layoutPath, lines, out layout)) return false;

            XDocument document;
            using (var zip = ZipFile.OpenRead(outputPath))
            using (var stream = zip.GetEntry("word/document.xml").Open())
                document = XDocument.Load(stream, LoadOptions.PreserveWhitespace);
            // Build all formatting in memory before replacing the document part.
            // A failure to use optional geometry leaves the plain document intact.
            try { ApplyLayout(document, lines, layout); }
            catch (ArgumentException) { return false; }
            catch (InvalidOperationException) { return false; }
            catch (System.Runtime.InteropServices.ExternalException) { return false; }

            using (var zip = ZipFile.Open(outputPath, ZipArchiveMode.Update))
            {
                zip.GetEntry("word/document.xml").Delete();
                using (var stream = zip.CreateEntry("word/document.xml").Open())
                using (var writer = XmlWriter.Create(stream, new XmlWriterSettings { Encoding = new UTF8Encoding(false) }))
                    document.Save(writer);
            }
            return true;
        }

        private static bool TryReadLayout(string path, string[][] text, out LayoutData layout)
        {
            layout = null;
            if (String.IsNullOrEmpty(path)) return false;
            try
            {
                var file = new FileInfo(path);
                if (!file.Exists || file.Length <= 0 || file.Length > MaximumMetadataBytes) return false;
                var serializer = new JavaScriptSerializer { MaxJsonLength = MaximumMetadataBytes, RecursionLimit = 32 };
                LayoutData candidate = serializer.Deserialize<LayoutData>(File.ReadAllText(path, Encoding.UTF8));
                if (candidate == null || candidate.Version != 1 || candidate.Pages == null || candidate.Pages.Length != text.Length) return false;
                for (int index = 0; index < candidate.Pages.Length; index++)
                {
                    PageData page = candidate.Pages[index];
                    // Word supports pages up to 22 inches in either direction.
                    // Very large PDF sheets keep their text but use the plain fallback.
                    if (page == null || !Bounded(page.Width, 96, 2112) || !Bounded(page.Height, 96, 2112) ||
                        !Bounded(page.ImageWidth, 1, 20000) || !Bounded(page.ImageHeight, 1, 20000) ||
                        page.Lines == null || page.Lines.Length != text[index].Length) return false;
                    foreach (LineData line in page.Lines)
                    {
                        if (line == null || !Bounded(line.X, 0, page.ImageWidth.Value) || !Bounded(line.Y, 0, page.ImageHeight.Value) ||
                            !Bounded(line.Width, 0.01, page.ImageWidth.Value) || !Bounded(line.Height, 0.01, page.ImageHeight.Value) ||
                            line.X.Value + line.Width.Value > page.ImageWidth.Value + 1 ||
                            line.Y.Value + line.Height.Value > page.ImageHeight.Value + 1) return false;
                    }
                }
                layout = candidate;
                return true;
            }
            catch (IOException) { return false; }
            catch (UnauthorizedAccessException) { return false; }
            catch (ArgumentException) { return false; }
            catch (InvalidOperationException) { return false; }
            catch (NotSupportedException) { return false; }
            catch (FormatException) { return false; }
            catch (OverflowException) { return false; }
        }

        private static bool Bounded(double? value, double minimum, double maximum)
        {
            return value.HasValue && !Double.IsNaN(value.Value) && !Double.IsInfinity(value.Value) && value.Value >= minimum && value.Value <= maximum;
        }

        private static void ApplyLayout(XDocument document, string[][] lines, LayoutData layout)
        {
            XElement body = document.Root.Element(W + "body");
            XElement[] paragraphs = body.Elements(W + "p").ToArray();
            if (paragraphs.Length != lines.Sum(page => page.Length) + lines.Length - 1)
                throw new InvalidOperationException("Recognised paragraph structure differed.");
            body.Elements(W + "sectPr").Remove();
            int paragraphIndex = 0;
            using (var bitmap = new Bitmap(1, 1))
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                for (int pageIndex = 0; pageIndex < layout.Pages.Length; pageIndex++)
                {
                    PageData page = layout.Pages[pageIndex];
                    var geometry = new PageGeometry(page);
                    var fontSizes = page.Lines.Select(line => EstimateFontSize(line.Height.Value * geometry.VerticalPoints)).ToArray();
                    double medianFont = fontSizes.OrderBy(value => value).ElementAt(fontSizes.Length / 2);
                    double previousFont = medianFont;
                    for (int index = 0; index < page.Lines.Length; index++)
                    {
                        LineData line = page.Lines[index];
                        double font = fontSizes[index];
                        if (Math.Abs(font - medianFont) <= medianFont * 0.12) font = medianFont;
                        font = FitFont(graphics, lines[pageIndex][index], font, line.Width.Value * geometry.HorizontalPoints);
                        double top = line.Y.Value * geometry.VerticalPoints;
                        double before = index == 0 ? Math.Max(0, top - geometry.Top / 20.0) :
                            Math.Max(0, top - page.Lines[index - 1].Y.Value * geometry.VerticalPoints - previousFont * 1.10);
                        // A column transition can move upwards or jump sideways. Keep
                        // the established reading order and avoid inventing a large gap.
                        if (index > 0 && Math.Abs(line.X.Value - page.Lines[index - 1].X.Value) > page.ImageWidth.Value * 0.25) before = 0;
                        if (index > 0) before = Math.Min(before, 72);
                        double left = Math.Max(0, line.X.Value * geometry.HorizontalPoints - geometry.Left / 20.0);
                        XElement paragraph = paragraphs[paragraphIndex++];
                        paragraph.Elements(W + "pPr").Remove();
                        paragraph.AddFirst(new XElement(W + "pPr",
                            new XElement(W + "widowControl", new XAttribute(W + "val", "0")),
                            new XElement(W + "snapToGrid", new XAttribute(W + "val", "0")),
                            new XElement(W + "spacing", new XAttribute(W + "before", Twips(before)), new XAttribute(W + "after", "0"),
                                new XAttribute(W + "line", Twips(font * 1.10)), new XAttribute(W + "lineRule", "atLeast")),
                            new XElement(W + "ind", new XAttribute(W + "left", Twips(left)), new XAttribute(W + "right", "0"))));
                        foreach (XElement run in paragraph.Elements(W + "r"))
                        {
                            run.Elements(W + "rPr").Remove();
                            run.AddFirst(FontProperties(font));
                        }
                        previousFont = font;
                    }
                    XElement section = geometry.Section();
                    if (pageIndex == layout.Pages.Length - 1) body.Add(section);
                    else
                    {
                        // The existing page separator stays a whitespace-only
                        // paragraph. A next-page section retains mixed paper sizes;
                        // turning its page break into a line break avoids two breaks.
                        XElement separator = paragraphs[paragraphIndex++];
                        separator.AddFirst(new XElement(W + "pPr",
                            new XElement(W + "spacing", new XAttribute(W + "before", "0"), new XAttribute(W + "after", "0"),
                                new XAttribute(W + "line", "20"), new XAttribute(W + "lineRule", "exact")), section));
                        foreach (XElement run in separator.Elements(W + "r")) run.AddFirst(FontProperties(1));
                        foreach (XElement pageBreak in separator.Descendants(W + "br")) pageBreak.Attribute(W + "type").Remove();
                    }
                }
            }
        }

        private static XElement FontProperties(double points)
        {
            int halfPoints = Math.Max(2, (int)Math.Round(points * 2));
            return new XElement(W + "rPr",
                new XElement(W + "rFonts", new XAttribute(W + "ascii", "Arial"), new XAttribute(W + "hAnsi", "Arial"), new XAttribute(W + "cs", "Arial")),
                new XElement(W + "sz", new XAttribute(W + "val", halfPoints)),
                new XElement(W + "szCs", new XAttribute(W + "val", halfPoints)));
        }

        private static double EstimateFontSize(double glyphHeight)
        {
            return Math.Round(Math.Max(6, Math.Min(36, glyphHeight * 1.25)) * 2) / 2;
        }

        private static double FitFont(Graphics graphics, string text, double points, double targetWidth)
        {
            if (String.IsNullOrWhiteSpace(text) || text.Length > 10000 || targetWidth <= 0) return points;
            using (var font = new Font("Arial", (float)points, FontStyle.Regular, GraphicsUnit.Point))
            using (var format = (StringFormat)StringFormat.GenericTypographic.Clone())
            {
                format.FormatFlags |= StringFormatFlags.MeasureTrailingSpaces;
                double measured = graphics.MeasureString(text, font, Int32.MaxValue, format).Width * 72.0 / graphics.DpiX;
                if (measured > targetWidth + 2) points *= (targetWidth + 2) / measured;
            }
            return Math.Round(Math.Max(6, points) * 2) / 2;
        }

        private static int Twips(double points) { return (int)Math.Round(points * 20); }

        private sealed class PageGeometry
        {
            public readonly double HorizontalPoints, VerticalPoints;
            public readonly int Width, Height, Left, Right, Top, Bottom;
            public PageGeometry(PageData page)
            {
                HorizontalPoints = page.Width.Value * 0.75 / page.ImageWidth.Value;
                VerticalPoints = page.Height.Value * 0.75 / page.ImageHeight.Value;
                Width = Twips(page.Width.Value * 0.75); Height = Twips(page.Height.Value * 0.75);
                Left = Math.Min(1440, Twips(page.Lines.Min(line => line.X.Value) * HorizontalPoints));
                Right = Math.Max(0, Math.Min(1440, Width - Twips(page.Lines.Max(line => line.X.Value + line.Width.Value) * HorizontalPoints)));
                Top = Math.Min(1440, Twips(page.Lines.Min(line => line.Y.Value) * VerticalPoints));
                Bottom = Math.Max(0, Math.Min(1440, Height - Twips(page.Lines.Max(line => line.Y.Value + line.Height.Value) * VerticalPoints)));
                // Leave a usable text region even for a sparse central caption.
                Left = Math.Min(Left, Width / 4); Right = Math.Min(Right, Width / 4);
                Top = Math.Min(Top, Height / 4); Bottom = Math.Min(Bottom, Height / 4);
            }
            public XElement Section()
            {
                var size = new XElement(W + "pgSz", new XAttribute(W + "w", Width), new XAttribute(W + "h", Height));
                if (Width > Height) size.Add(new XAttribute(W + "orient", "landscape"));
                return new XElement(W + "sectPr", new XElement(W + "type", new XAttribute(W + "val", "nextPage")), size,
                    new XElement(W + "pgMar", new XAttribute(W + "top", Top), new XAttribute(W + "right", Right), new XAttribute(W + "bottom", Bottom),
                        new XAttribute(W + "left", Left), new XAttribute(W + "header", "0"), new XAttribute(W + "footer", "0"), new XAttribute(W + "gutter", "0")));
            }
        }

        private sealed class LayoutData
        {
            public LayoutData() { }
            public int Version { get; set; }
            public PageData[] Pages { get; set; }
        }
        private sealed class PageData
        {
            public PageData() { }
            public double? Width { get; set; }
            public double? Height { get; set; }
            public double? ImageWidth { get; set; }
            public double? ImageHeight { get; set; }
            public LineData[] Lines { get; set; }
        }
        private sealed class LineData
        {
            public LineData() { }
            public double? X { get; set; }
            public double? Y { get; set; }
            public double? Width { get; set; }
            public double? Height { get; set; }
        }
    }
}
