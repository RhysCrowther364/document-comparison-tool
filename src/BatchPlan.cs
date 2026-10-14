using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;

namespace LocalDocumentCompare
{
    public sealed class BatchInput
    {
        public string Path { get; set; }
        public bool UseOcr { get; set; }
    }

    public sealed class BatchPair
    {
        public int Number { get; internal set; }
        public int ArchiveIndex { get; internal set; }
        public int NewIndex { get; internal set; }
        public BatchInput Archive { get; internal set; }
        public BatchInput New { get; internal set; }
        public string OutputPath { get; internal set; }
    }

    public sealed class BatchPlan
    {
        public IList<BatchInput> Inputs { get; private set; }
        public IList<BatchPair> Pairs { get; private set; }
        public string OutputDirectory { get; private set; }
        public string JobType { get; private set; }
        public bool IsMultiArchive { get; private set; }
        public string ReportPath { get { return Path.Combine(OutputDirectory, IsMultiArchive ? "multi-archive-results.csv" : "batch-results.csv"); } }
        internal readonly object ReportLock = new object();
        internal bool OwnsReport;

        private BatchPlan() { }

        public static List<BatchInput> LoadFolder(string folder)
        {
            ComparisonEngine.RequireSourcePath(folder);
            string full = Path.GetFullPath(folder);
            if (!Directory.Exists(full)) throw new InvalidOperationException("The input folder is unavailable. Check the path, permissions and any shared-drive connection.");
            var paths = Directory.EnumerateFiles(full, "*", SearchOption.TopDirectoryOnly)
                .Where(path => Supported(path) && !Path.GetFileName(path).StartsWith("~$", StringComparison.Ordinal))
                .ToList();
            paths.Sort(new NaturalFilenameComparer());
            return paths.Select(path => new BatchInput { Path = Path.GetFullPath(path), UseOcr = false }).ToList();
        }

        public static BatchPlan Create(IList<BatchInput> ordered, string outputRoot, string jobType)
        {
            if (ordered == null || ordered.Count < 2)
                throw new InvalidOperationException("Choose at least two documents for an all-pairs comparison.");
            return CreateCore(ordered, outputRoot, jobType, false);
        }

        public static BatchPlan CreateMultiArchive(IList<BatchInput> archives, BatchInput newDocument, string outputRoot, string jobType)
        {
            if (archives == null || archives.Count == 0)
                throw new InvalidOperationException("Choose at least one archive document.");
            if (newDocument == null) throw new InvalidOperationException("Choose the new document to assess.");
            var inputs = new List<BatchInput>(archives);
            inputs.Add(newDocument);
            return CreateCore(inputs, outputRoot, jobType, true);
        }

        private static BatchPlan CreateCore(IList<BatchInput> ordered, string outputRoot, string jobType, bool multiArchive)
        {
            string type = (jobType ?? "").Trim().ToUpperInvariant();
            if (type != "WF" && type != "EP") throw new InvalidOperationException("Select WF or EP as the batch job type.");
            ComparisonEngine.RequireOutputPath(outputRoot);
            string root = Path.GetFullPath(outputRoot);
            if (File.Exists(root)) throw new InvalidOperationException("Choose an output folder, not a file.");

            var inputs = new List<BatchInput>();
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (BatchInput input in ordered)
            {
                if (input == null) throw new InvalidOperationException("The document list contains an empty entry.");
                ComparisonEngine.RequireSourcePath(input.Path);
                string path = Path.GetFullPath(input.Path);
                if (!File.Exists(path)) throw new InvalidOperationException("A batch input file could not be found or read. Check the document list, permissions and any shared-drive connection.");
                if (!Supported(path) || Path.GetFileName(path).StartsWith("~$", StringComparison.Ordinal))
                    throw new InvalidOperationException("Batch inputs must be PDF, DOC or DOCX documents. Word lock files cannot be compared.");
                if (!seen.Add(path)) throw new InvalidOperationException("Each batch input must be a different file. Remove the duplicate entry.");
                if (input.UseOcr && !String.Equals(Path.GetExtension(path), ".pdf", StringComparison.OrdinalIgnoreCase))
                    throw new InvalidOperationException("Local OCR can only be selected for PDF inputs.");
                // Snapshot the selection so later list edits cannot reorder a running plan.
                inputs.Add(new BatchInput { Path = path, UseOcr = input.UseOcr });
            }
            long count = multiArchive ? inputs.Count - 1 : (long)inputs.Count * (inputs.Count - 1) / 2;
            if (count > Int32.MaxValue) throw new InvalidOperationException("There are too many files to represent every pair in this application.");
            string output;
            do
            {
                output = Path.Combine(root, (multiArchive ? "Multi Archive " : "Batch ") + DateTime.Now.ToString("yyyyMMdd-HHmmss", CultureInfo.InvariantCulture) + "-" + Guid.NewGuid().ToString("N").Substring(0, 8));
            }
            while (Directory.Exists(output) || File.Exists(output));

            var pairs = new List<BatchPair>();
            int digits = Math.Max(3, count.ToString(CultureInfo.InvariantCulture).Length);
            for (int archive = 0; archive < inputs.Count - 1; archive++)
            {
                for (int revised = multiArchive ? inputs.Count - 1 : archive + 1; revised < inputs.Count; revised++)
                {
                    int number = pairs.Count + 1;
                    string folder = Path.Combine(output, number.ToString("D" + digits, CultureInfo.InvariantCulture));
                    string filename = Path.GetFileNameWithoutExtension(inputs[revised].Path) + " vs " + Path.GetFileNameWithoutExtension(inputs[archive].Path) + ".docx";
                    string path = Path.Combine(folder, filename);
                    // Match the single-comparison limit for Word/.NET Framework.
                    // Check every prospective path before creating output folders.
                    if (path.Length > 240 || filename.Length > 255)
                        throw new InvalidOperationException("The output path for batch pair " + number + " is too long. Shorten the source filenames or choose a shorter output folder.");
                    pairs.Add(new BatchPair { Number = number, ArchiveIndex = archive, NewIndex = revised,
                        Archive = inputs[archive], New = inputs[revised], OutputPath = path });
                }
            }
            if (Path.Combine(output, multiArchive ? "multi-archive-results.csv" : "batch-results.csv").Length > 240)
                throw new InvalidOperationException("The batch output folder path is too long. Choose a shorter output folder.");

            var created = new List<string>();
            try
            {
                Directory.CreateDirectory(root);
                if (Directory.Exists(output) || File.Exists(output))
                    throw new IOException("The batch output name was taken by another operation. Start the batch again.");
                Directory.CreateDirectory(output);
                created.Add(output);
                foreach (BatchPair pair in pairs)
                {
                    string folder = Path.GetDirectoryName(pair.OutputPath);
                    Directory.CreateDirectory(folder);
                    created.Add(folder);
                }
            }
            catch
            {
                // Remove only empty directories created by this attempt. Never recurse
                // through a result folder another operation might have populated.
                for (int i = created.Count - 1; i >= 0; i--)
                    try { Directory.Delete(created[i], false); } catch { }
                throw;
            }
            return new BatchPlan { Inputs = new ReadOnlyCollection<BatchInput>(inputs), Pairs = new ReadOnlyCollection<BatchPair>(pairs),
                OutputDirectory = output, JobType = type, IsMultiArchive = multiArchive };
        }

        private static bool Supported(string path)
        {
            string extension = Path.GetExtension(path);
            return String.Equals(extension, ".docx", StringComparison.OrdinalIgnoreCase) ||
                String.Equals(extension, ".doc", StringComparison.OrdinalIgnoreCase) ||
                String.Equals(extension, ".pdf", StringComparison.OrdinalIgnoreCase);
        }

        private sealed class NaturalFilenameComparer : IComparer<string>
        {
            public int Compare(string leftPath, string rightPath)
            {
                string left = Path.GetFileName(leftPath), right = Path.GetFileName(rightPath);
                int a = 0, b = 0;
                while (a < left.Length && b < right.Length)
                {
                    if (Digit(left[a]) && Digit(right[b]))
                    {
                        int endA = a, endB = b;
                        while (endA < left.Length && Digit(left[endA])) endA++;
                        while (endB < right.Length && Digit(right[endB])) endB++;
                        while (a < endA - 1 && left[a] == '0') a++;
                        while (b < endB - 1 && right[b] == '0') b++;
                        int comparison = (endA - a).CompareTo(endB - b);
                        if (comparison != 0) return comparison;
                        comparison = String.Compare(left, a, right, b, endA - a, StringComparison.Ordinal);
                        if (comparison != 0) return comparison;
                        a = endA; b = endB;
                    }
                    else
                    {
                        int comparison = Char.ToUpperInvariant(left[a]).CompareTo(Char.ToUpperInvariant(right[b]));
                        if (comparison != 0) return comparison;
                        a++; b++;
                    }
                }
                int remaining = (left.Length - a).CompareTo(right.Length - b);
                return remaining != 0 ? remaining : StringComparer.Ordinal.Compare(left, right);
            }
            private static bool Digit(char value) { return value >= '0' && value <= '9'; }
        }
    }

    public sealed class BatchRecord
    {
        public BatchPair Pair { get; set; }
        public string State { get; set; }
        public ComparisonResult Result { get; set; }
        public BatchRecord() { State = "Pending"; }
    }

    public static class BatchReport
    {
        private static readonly HashSet<string> States = new HashSet<string>(new[] {
            "Pending", "Running", "Review required", "Completed", "Failed", "Cancelled", "Not run"
        }, StringComparer.Ordinal);

        public static string Save(BatchPlan plan, IList<BatchRecord> records)
        {
            if (plan == null) throw new ArgumentNullException("plan");
            if (records == null) throw new ArgumentNullException("records");
            lock (plan.ReportLock)
            {
                var indexed = new Dictionary<int, BatchRecord>();
                foreach (BatchRecord record in records)
                {
                    if (record == null || record.Pair == null || record.Pair.Number < 1 || record.Pair.Number > plan.Pairs.Count)
                        throw new ArgumentException("Every batch record must identify a pair in this plan.", "records");
                    BatchPair pair = plan.Pairs[record.Pair.Number - 1];
                    if (record.Pair.ArchiveIndex != pair.ArchiveIndex || record.Pair.NewIndex != pair.NewIndex ||
                        !String.Equals(record.Pair.OutputPath, pair.OutputPath, StringComparison.OrdinalIgnoreCase))
                        throw new ArgumentException("A batch record belongs to another plan.", "records");
                    if (!States.Contains(record.State ?? "")) throw new ArgumentException("A batch record has an unknown state.", "records");
                    if (indexed.ContainsKey(pair.Number)) throw new ArgumentException("A batch pair has duplicate report records.", "records");
                    indexed.Add(pair.Number, record);
                }
                string report = plan.ReportPath;
                if (File.Exists(report) && !plan.OwnsReport)
                    throw new IOException("A batch results report already exists at this path. It was not created by this batch and will not be overwritten.");
                // Keep the staging filename shorter than batch-results.csv so
                // it also fits every output path accepted by the planner.
                string temporary = Path.Combine(plan.OutputDirectory, ".r" + Guid.NewGuid().ToString("N").Substring(0, 10) + ".tmp");
                try
                {
                    using (var stream = new FileStream(temporary, FileMode.CreateNew, FileAccess.Write, FileShare.None))
                    using (var writer = new StreamWriter(stream, new UTF8Encoding(true)))
                    {
                        writer.NewLine = "\r\n";
                        var headers = new List<string>(new[] { "Pair", "Archive", "New", "Scope", "Status", "Crossover (%)", "New file word count", "Highlighted insertion word count", "Comparison document", "Notes" });
                        if (plan.IsMultiArchive) headers.Insert(0, "Rank");
                        Row(writer, headers);
                        var all = plan.Pairs.Select(pair => indexed.ContainsKey(pair.Number) ? indexed[pair.Number] : new BatchRecord { Pair = pair }).ToList();
                        Dictionary<int, int> ranks = plan.IsMultiArchive ? MultiArchiveRanking.Ranks(all, plan.JobType) : null;
                        foreach (BatchRecord record in plan.IsMultiArchive ? MultiArchiveRanking.Ordered(all, plan.JobType) : all)
                        {
                            int rank;
                            WriteRecord(writer, plan, record, ranks != null && ranks.TryGetValue(record.Pair.Number, out rank) ? (int?)rank : null);
                        }
                    }
                    if (File.Exists(report))
                    {
                        if (!plan.OwnsReport) throw new IOException("Another operation created this report. It will not be overwritten.");
                        // Ignore failures to merge optional metadata/ACLs; replacement
                        // itself remains atomic and still fails if the file is locked.
                        File.Replace(temporary, report, null, true);
                    }
                    else File.Move(temporary, report);
                    plan.OwnsReport = true;
                    return report;
                }
                finally { try { if (File.Exists(temporary)) File.Delete(temporary); } catch { } }
            }
        }

        private static void WriteRecord(TextWriter writer, BatchPlan plan, BatchRecord record, int? rank)
        {
            ComparisonResult result = record.Result;
            string state = record.State;
            var notes = new List<string>();
            if (result != null)
            {
                if (!String.IsNullOrWhiteSpace(result.Error)) notes.Add(result.Error);
                if (result.Warnings != null) notes.AddRange(result.Warnings.Where(w => !String.IsNullOrWhiteSpace(w)));
            }
            if (state == "Completed" && result != null && result.LanguageReviewRequired)
                state = "Review required";
            if (state == "Completed" && (result == null || !result.Success))
            {
                state = "Failed";
                notes.Add("No successful comparison result was returned.");
            }
            ComparisonSectionResult[] sections = null;
            if (state == "Completed" && plan.JobType == "EP")
            {
                string[] scopes = { "Description", "Full specification", "Claims" };
                if (result.Sections == null || scopes.Any(scope => result.Sections.Count(s => s != null && String.Equals(s.Scope, scope, StringComparison.OrdinalIgnoreCase)) != 1))
                {
                    state = "Failed";
                    notes.Add("The EP result did not contain all three required section figures.");
                }
                else sections = scopes.Select(scope => result.Sections.Single(s => s != null && String.Equals(s.Scope, scope, StringComparison.OrdinalIgnoreCase))).ToArray();
            }
            string noteText = String.Join(" | ", notes.Distinct(StringComparer.Ordinal));
            if (state != "Completed") WriteResultRow(writer, plan, record.Pair, "", state, null, "", noteText, null);
            else if (sections != null)
                foreach (ComparisonSectionResult section in sections)
                    WriteResultRow(writer, plan, record.Pair, section.Scope, state, section, RelativeOutput(plan, record.Pair), noteText, rank);
            else
                WriteResultRow(writer, plan, record.Pair, "Document", state,
                    new ComparisonSectionResult { NewWords = result.NewWords, InsertedWords = result.InsertedWords, CrossoverPercent = result.CrossoverPercent },
                    RelativeOutput(plan, record.Pair), noteText, rank);
        }

        private static string RelativeOutput(BatchPlan plan, BatchPair pair)
        {
            string prefix = plan.OutputDirectory.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar) + Path.DirectorySeparatorChar;
            if (!pair.OutputPath.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("A comparison output is outside this batch folder.");
            return pair.OutputPath.Substring(prefix.Length);
        }

        private static void WriteResultRow(TextWriter writer, BatchPlan plan, BatchPair pair, string scope, string state, ComparisonSectionResult figures, string document, string notes, int? rank)
        {
            string crossover = "";
            if (figures != null && figures.CrossoverPercent.HasValue && !Double.IsNaN(figures.CrossoverPercent.Value) && !Double.IsInfinity(figures.CrossoverPercent.Value))
                crossover = figures.CrossoverPercent.Value.ToString("0.0", CultureInfo.InvariantCulture);
            var fields = new List<string>(new[] {
                pair.Number.ToString(CultureInfo.InvariantCulture), Path.GetFileName(pair.Archive.Path), Path.GetFileName(pair.New.Path), scope, state, crossover,
                figures == null ? "" : figures.NewWords.ToString(CultureInfo.InvariantCulture),
                figures == null ? "" : figures.InsertedWords.ToString(CultureInfo.InvariantCulture), document, notes
            });
            if (plan.IsMultiArchive) fields.Insert(0, rank.HasValue ? rank.Value.ToString(CultureInfo.InvariantCulture) : "");
            Row(writer, fields);
        }

        private static void Row(TextWriter writer, IEnumerable<string> fields)
        {
            writer.WriteLine(String.Join(",", fields.Select(CsvCell)));
        }

        private static string CsvCell(string value)
        {
            value = value ?? "";
            int first = 0;
            while (first < value.Length && (Char.IsWhiteSpace(value[first]) || Char.IsControl(value[first]) ||
                Char.GetUnicodeCategory(value[first]) == UnicodeCategory.Format)) first++;
            // CSV quoting alone does not stop Excel from treating an imported cell
            // as a formula. Prefix dangerous user-controlled text with an apostrophe.
            if (first < value.Length && "=+-@".IndexOf(value[first]) >= 0) value = "'" + value;
            return "\"" + value.Replace("\"", "\"\"") + "\"";
        }
    }
}
