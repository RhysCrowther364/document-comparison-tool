using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Xml.Linq;
using LocalDocumentCompare;

// Synthetic, local preflight cases: these stop before starting Microsoft Word.
internal static class LanguageMismatchNoticeTests
{
    private const string English = "The present invention relates to a method for controlling a motor. The apparatus comprises a housing, a controller and a sensor. The controller receives the signal from the sensor and adjusts the speed of the motor. This arrangement improves the operation of the device. ";
    private const string German = "Die vorliegende Erfindung betrifft ein Verfahren zur Steuerung eines Motors. Die Vorrichtung umfasst ein Gehäuse, eine Steuereinheit und einen Sensor. Die Steuereinheit empfängt das Signal des Sensors und stellt die Drehzahl des Motors ein. Diese Anordnung verbessert den Betrieb der Vorrichtung. ";
    private static int checks;
    private static string root;

    public static int Main()
    {
        string parent = Path.GetFullPath(Path.Combine(Path.GetTempPath(), "LocalDocumentCompare")) + Path.DirectorySeparatorChar;
        root = Path.Combine(parent, "language-notice-tests-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        try
        {
            string en = Repeat(English), de = Repeat(German);
            string archive = Docx("Archive-German", de), revised = Docx("New-English", en);
            ComparisonResult mismatch = Run("wf", archive, revised, "WF", false);
            CheckMismatch(mismatch, "Document", "German", "English", "WF mismatch");
            CheckMismatch(Run("wf-confirmed", archive, revised, "WF", true), "Document", "German", "English", "confirmed mismatch remains blocked");
            string epArchive = Docx("EP-German", "Beschreibung", de, "Ansprüche", "1. " + de);
            string epNew = Docx("EP-English", "Description", en, "Claims", "1. " + en);
            CheckMismatch(Run("ep-description", epArchive, epNew, "EP", false), "Description", "German", "English", "EP description mismatch");
            string epSameDescription = Docx("EP-English-Description-German-Claims", "Description", en, "Ansprüche", "1. " + de);
            CheckMismatch(Run("ep-claims", epSameDescription, epNew, "EP", true), "Claims", "German", "English", "EP claims mismatch");
            string shortDoc = Docx("Short", "A motor comprises a sensor.");
            ComparisonResult shortReview = Run("short", shortDoc, shortDoc, "WF", false);
            Equal(true, shortReview.LanguageReviewRequired, "short prose retains manual review");
            Equal(null, shortReview.LanguageMismatch, "manual review is not classified as known mismatch");
            Equal(null, LanguageMismatchNotice.SingleMessage(shortReview), "manual review has no mismatch popup");
            string mixedDoc = Docx("Mixed", en, de);
            ComparisonResult mixedReview = Run("mixed", mixedDoc, mixedDoc, "WF", false);
            Equal(true, mixedReview.LanguageReviewRequired, "mixed prose retains manual review");
            Equal(null, mixedReview.LanguageMismatch, "mixed text has no known mismatch classification");
            Equal(null, LanguageMismatchNotice.SingleMessage(mixedReview), "mixed review has no mismatch popup");

            string roundtrip = Path.Combine(root, "result.json"); JsonFile.Write(roundtrip, mismatch);
            ComparisonResult restored = JsonFile.Read<ComparisonResult>(roundtrip);
            CheckMismatch(restored, "Document", "German", "English", "worker JSON round trip");
            Equal(null, LanguageMismatchNotice.SingleMessage(null), "no result has no notice");
            Equal(null, LanguageMismatchNotice.SingleMessage(new ComparisonResult { Success = false, Error = "Source language mismatch appears in an unrelated filename." }), "error text is not parsed as a mismatch");
            Equal(null, LanguageMismatchNotice.SingleMessage(new ComparisonResult { Success = true, LanguageMismatch = mismatch.LanguageMismatch }), "successful result never shows stale mismatch");

            var records = new List<BatchRecord>();
            records.Add(Record(1, "Failed", mismatch));
            records.Add(Record(2, "Completed", new ComparisonResult { Success = true }));
            records.Add(Record(3, "Failed", new ComparisonResult { Error = "Other failure" }));
            string batch = LanguageMismatchNotice.BatchMessage(records);
            Equal(true, batch != null && batch.Contains("01 new.docx") && batch.Contains("01 archive.docx"), "batch identifies failed filenames");
            Equal(true, batch.Contains("German") && batch.Contains("English"), "batch gives the detected languages");
            Equal(false, batch.Contains("02 new.docx") || batch.Contains("03 new.docx"), "batch excludes successful and unrelated failures");
            Equal(null, LanguageMismatchNotice.BatchMessage(new List<BatchRecord> { Record(1, "Cancelled", mismatch) }), "cancelled pair does not show stale mismatch");
            Equal(null, LanguageMismatchNotice.BatchMessage(new List<BatchRecord> { Record(1, "Review required", shortReview) }), "review pause does not show mismatch summary");
            Equal(null, LanguageMismatchNotice.BatchMessage(new List<BatchRecord>()), "empty batch has no popup");
            var many = new List<BatchRecord>();
            for (int index = 1; index <= 14; index++) many.Add(Record(index, "Failed", mismatch));
            string bounded = LanguageMismatchNotice.BatchMessage(many);
            Equal(true, bounded != null && bounded.Contains("05 new.docx"), "batch summary includes five detailed pairs");
            Equal(false, bounded.Contains("06 new.docx") || bounded.Contains("14 new.docx"), "batch summary remains bounded for large jobs");
            Equal(true, bounded.Length < 5000, "large summary stays short enough for a dialog");
            Console.WriteLine("PASS: " + checks + " language notice/Windows ELS checks; known mismatches stop before Word; sources unchanged; no output comparison published.");
            return 0;
        }
        catch (Exception error) { Console.Error.WriteLine(error); return 1; }
        finally
        {
            string resolved = Path.GetFullPath(root);
            if (resolved.StartsWith(parent, StringComparison.OrdinalIgnoreCase) && Path.GetFileName(resolved).StartsWith("language-notice-tests-", StringComparison.Ordinal))
                try { Directory.Delete(resolved, true); } catch { }
        }
    }

    private static BatchRecord Record(int number, string state, ComparisonResult result)
    {
        return new BatchRecord { State = state, Result = result, Pair = new BatchPair { Number = number,
            Archive = new BatchInput { Path = Path.Combine(root, number.ToString("00") + " archive.docx") },
            New = new BatchInput { Path = Path.Combine(root, number.ToString("00") + " new.docx") } } };
    }
    private static void CheckMismatch(ComparisonResult result, string scope, string archive, string revised, string label)
    {
        Equal(false, result.Success, label + " fails");
        Equal(false, result.LanguageReviewRequired, label + " cannot be manually overridden");
        Equal(true, result.LanguageMismatch != null, label + " carries typed reason");
        Equal(scope, result.LanguageMismatch.Scope, label + " scope");
        Equal(archive, result.LanguageMismatch.ArchiveLanguage, label + " archive language");
        Equal(revised, result.LanguageMismatch.NewLanguage, label + " new language");
        string message = LanguageMismatchNotice.SingleMessage(result);
        Equal(true, message != null && message.Contains(archive) && message.Contains(revised), label + " clear languages in notice");
        Equal(true, message.Contains("same") && message.Contains("language"), label + " gives corrective action");
        Equal(null, result.OutputPath, label + " no advertised comparison");
        Equal(null, result.SummaryPath, label + " no advertised summary");
    }
    private static ComparisonResult Run(string name, string archive, string revised, string type, bool confirmed)
    {
        string work = Path.Combine(root, name); Directory.CreateDirectory(work);
        string output = Path.Combine(root, name + " comparison.docx");
        string archiveHash = Hash(archive), newHash = Hash(revised);
        ComparisonResult result = ComparisonEngine.Run(new ComparisonRequest { ArchivePath = archive, NewPath = revised,
            JobType = type, MatchingLanguageConfirmed = confirmed, WorkDirectory = work, OutputPath = output }, delegate { });
        Equal(archiveHash, Hash(archive), name + " unchanged archive");
        Equal(newHash, Hash(revised), name + " unchanged new source");
        Equal(false, File.Exists(Path.Combine(work, "word.process")), name + " never starts Word");
        Equal(false, File.Exists(output) || File.Exists(Path.ChangeExtension(output, ".txt")), name + " publishes no output");
        return result;
    }
    private static string Repeat(string value) { return String.Concat(Enumerable.Repeat(value, 14)); }
    private static string Docx(string name, params string[] paragraphs)
    {
        string path = Path.Combine(root, name + ".docx");
        XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        var document = new XDocument(new XElement(w + "document", new XAttribute(XNamespace.Xmlns + "w", w),
            new XElement(w + "body", paragraphs.Select(p => new XElement(w + "p", new XElement(w + "r", new XElement(w + "t", p)))), new XElement(w + "sectPr"))));
        using (ZipArchive zip = ZipFile.Open(path, ZipArchiveMode.Create))
        {
            Add(zip, "word/document.xml", document.ToString());
            Add(zip, "[Content_Types].xml", "<Types xmlns='http://schemas.openxmlformats.org/package/2006/content-types'><Default Extension='rels' ContentType='application/vnd.openxmlformats-package.relationships+xml'/><Default Extension='xml' ContentType='application/xml'/><Override PartName='/word/document.xml' ContentType='application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml'/></Types>");
            Add(zip, "_rels/.rels", "<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'><Relationship Id='rId1' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument' Target='word/document.xml'/></Relationships>");
        }
        return path;
    }
    private static void Add(ZipArchive zip, string name, string text) { using (var writer = new StreamWriter(zip.CreateEntry(name).Open(), new UTF8Encoding(false))) writer.Write(text); }
    private static string Hash(string path) { using (var sha = SHA256.Create()) using (var stream = File.OpenRead(path)) return BitConverter.ToString(sha.ComputeHash(stream)); }
    private static void Equal<T>(T expected, T actual, string label)
    {
        if (!EqualityComparer<T>.Default.Equals(expected, actual)) throw new Exception(label + ": expected " + expected + "; actual " + actual);
        checks++;
    }
}
