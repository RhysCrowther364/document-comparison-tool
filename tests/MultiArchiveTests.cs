using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using LocalDocumentCompare;

internal static class MultiArchiveTests
{
    private static int checks;
    public static int Main()
    {
        string root = Path.Combine(Path.GetTempPath(), "LocalCompare-MultiModel-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        try
        {
            var archives = new List<BatchInput>();
            foreach (string name in new[] { "Archive1.docx", "Archive2.docx", "Archive3.pdf", "Archive4.doc" })
                archives.Add(Input(root, name));
            archives[2].UseOcr = true;
            BatchInput revised = Input(root, "New.pdf"); revised.UseOcr = true;
            string destination = Path.Combine(root, "results");
            BatchPlan plan = BatchPlan.CreateMultiArchive(archives, revised, destination, "wf");
            Equal(true, plan.IsMultiArchive, "plan identifies one-to-many mode");
            Equal(4, plan.Pairs.Count, "four archives produce four comparisons");
            Equal(5, plan.Inputs.Count, "new document is separate from four archives");
            Equal("WF", plan.JobType, "mode normalized");
            Equal("multi-archive-results.csv", Path.GetFileName(plan.ReportPath), "dedicated ranked report");
            for (int i = 0; i < plan.Pairs.Count; i++)
            {
                BatchPair pair = plan.Pairs[i];
                Equal(i, pair.ArchiveIndex, "archive direction " + i);
                Equal(4, pair.NewIndex, "same new index " + i);
                Equal(revised.Path, pair.New.Path, "same new path " + i);
                Equal(true, pair.New.UseOcr, "new OCR snapshot " + i);
                Equal("New vs " + Path.GetFileNameWithoutExtension(archives[i].Path) + ".docx", Path.GetFileName(pair.OutputPath), "output naming " + i);
                Equal(true, Directory.Exists(Path.GetDirectoryName(pair.OutputPath)), "separate output folder " + i);
            }
            archives[0].Path = archives[1].Path; revised.UseOcr = false;
            Equal("Archive1.docx", Path.GetFileName(plan.Pairs[0].Archive.Path), "running archives isolated from later input edits");
            Equal(true, plan.Pairs[0].New.UseOcr, "running new settings isolated from edits");
            archives[0] = new BatchInput { Path = plan.Pairs[0].Archive.Path };

            var records = new List<BatchRecord> {
                Done(plan.Pairs[0], 10, 4, 60), Done(plan.Pairs[1], 10, 1, 90),
                Done(plan.Pairs[2], 10, 1, 90), Done(plan.Pairs[3], 10, 0, null)
            };
            Equal("2|3|1|4", String.Join("|", MultiArchiveRanking.Ordered(records, "WF").Select(r => r.Pair.Number)), "highest verified crossover first and unavailable last");
            Dictionary<int, int> ranks = MultiArchiveRanking.Ranks(records, "WF");
            Equal(1, ranks[2], "best archive rank"); Equal(1, ranks[3], "tied best archive shares rank");
            Equal(3, ranks[1], "next rank follows two joint best matches");
            Equal(false, ranks.ContainsKey(4), "counts alone cannot rank unavailable crossover");
            string csv = BatchReport.Save(plan, records);
            string[] lines = File.ReadAllLines(csv);
            Equal(5, lines.Length, "one WF report row per archive");
            Equal(true, lines[0].StartsWith("\"Rank\",\"Pair\""), "rank column in dedicated report");
            Equal(true, lines[1].StartsWith("\"1\",\"2\""), "CSV is ranked like UI");
            Equal(true, lines[2].StartsWith("\"1\",\"3\""), "CSV retains equal ranks");
            Equal(true, lines[4].StartsWith("\"\",\"4\""), "unavailable CSV result has no rank");
            using (var locked = new FileStream(csv, FileMode.Open, FileAccess.ReadWrite, FileShare.None))
                Reject(delegate { BatchReport.Save(plan, records); }, "locked report update refused");
            Equal(lines[1], File.ReadAllLines(csv)[1], "failed update preserves existing ranked report");
            Equal(csv, BatchReport.Save(plan, records), "report saving can be retried");

            BatchRecord invalid = Done(plan.Pairs[0], 10, 0, 100);
            foreach (string state in new[] { "Failed", "Cancelled", "Pending", "Running", "Review required", "Not run" })
            { invalid.State = state; Equal(null, MultiArchiveRanking.Score(invalid, "WF"), "incomplete state unranked: " + state); }
            invalid.State = "Completed";
            invalid.Result.Success = false; Equal(null, MultiArchiveRanking.Score(invalid, "WF"), "unsuccessful result unranked");
            invalid.Result.Success = true; invalid.Result.LanguageReviewRequired = true;
            Equal(null, MultiArchiveRanking.Score(invalid, "WF"), "language review cannot be a best match");
            invalid.Result.LanguageReviewRequired = false;
            foreach (double score in new[] { Double.NaN, Double.PositiveInfinity, -1.0, 101.0, 95.0 })
            { invalid.Result.CrossoverPercent = score; Equal(null, MultiArchiveRanking.Score(invalid, "WF"), "invalid or inconsistent percentage excluded"); }
            invalid.Result.CrossoverPercent = 0; invalid.Result.NewWords = 0;
            Equal(null, MultiArchiveRanking.Score(invalid, "WF"), "empty new document unranked");

            BatchPlan ep = BatchPlan.CreateMultiArchive(archives.Take(3).ToList(), revised, destination, "EP");
            var epRecords = new List<BatchRecord> { EpDone(ep.Pairs[0], 0, 2), EpDone(ep.Pairs[1], 1, 0), EpDone(ep.Pairs[2], 0, 0) };
            epRecords[2].Result.Sections[1].CrossoverPercent = null;
            epRecords[2].Result.CrossoverPercent = 100;
            Equal(80.0, MultiArchiveRanking.Score(epRecords[0], "EP"), "EP uses full specification instead of description");
            Equal(90.0, MultiArchiveRanking.Score(epRecords[1], "EP"), "EP full specification score");
            Equal(null, MultiArchiveRanking.Score(epRecords[2], "EP"), "EP unavailable full score cannot fall back to another scope");
            Equal("2|1|3", String.Join("|", MultiArchiveRanking.Ordered(epRecords, "EP").Select(r => r.Pair.Number)), "EP ordering by full specification");
            string[] epLines = File.ReadAllLines(BatchReport.Save(ep, epRecords));
            Equal(10, epLines.Length, "three EP scopes per archive in ranked CSV");
            Equal(true, epLines.Skip(1).Take(3).All(l => l.StartsWith("\"1\",\"2\"")), "same full-spec rank shown beside all three scopes");
            epRecords[0].Result.Sections = epRecords[0].Result.Sections.Take(2).ToArray();
            Equal(null, MultiArchiveRanking.Score(epRecords[0], "EP"), "incomplete EP scope set is not ranked");

            BatchPlan one = BatchPlan.CreateMultiArchive(archives.Take(1).ToList(), revised, destination, "WF");
            Equal(1, one.Pairs.Count, "one archive is a valid search");
            Reject(delegate { BatchPlan.CreateMultiArchive(new List<BatchInput>(), revised, destination, "WF"); }, "empty archive selection refused");
            Reject(delegate { BatchPlan.CreateMultiArchive(archives, null, destination, "WF"); }, "missing new document refused");
            Reject(delegate { BatchPlan.CreateMultiArchive(archives, archives[0], destination, "WF"); }, "new file cannot also be an archive");
            Reject(delegate { BatchPlan.CreateMultiArchive(new[] { archives[0], archives[0] }, revised, destination, "WF"); }, "duplicate archives refused");
            Reject(delegate { BatchPlan.CreateMultiArchive(new[] { new BatchInput { Path = archives[0].Path, UseOcr = true } }, revised, destination, "WF"); }, "Word archive OCR refused");
            BatchInput newWord = Input(root, "NewWord.docx"); newWord.UseOcr = true;
            Reject(delegate { BatchPlan.CreateMultiArchive(archives, newWord, destination, "WF"); }, "Word new OCR refused");
            Reject(delegate { BatchPlan.CreateMultiArchive(archives, new BatchInput { Path = Path.Combine(root, "missing.docx") }, destination, "WF"); }, "missing new file refused");
            Equal(false, BatchPlan.Create(archives.Take(2).ToList(), destination, "WF").IsMultiArchive, "ordinary all-pairs mode remains separate");
            Console.WriteLine("PASS: " + checks + " Multi Archive planning, ranking, report and validation checks.");
            return 0;
        }
        catch (Exception ex) { Console.Error.WriteLine(ex); return 1; }
        finally
        {
            string resolved = Path.GetFullPath(root);
            string boundary = Path.GetFullPath(Path.GetTempPath()).TrimEnd(Path.DirectorySeparatorChar) + Path.DirectorySeparatorChar;
            if (!resolved.StartsWith(boundary, StringComparison.OrdinalIgnoreCase) || !Path.GetFileName(resolved).StartsWith("LocalCompare-MultiModel-", StringComparison.Ordinal)) throw new Exception("Unexpected test cleanup path.");
            Directory.Delete(resolved, true);
        }
    }

    private static BatchInput Input(string root, string name) { string path = Path.Combine(root, name); File.WriteAllText(path, "synthetic placeholder; never opened in Word"); return new BatchInput { Path = path }; }
    private static BatchRecord Done(BatchPair pair, int words, int inserted, double? crossover)
    { return new BatchRecord { Pair = pair, State = "Completed", Result = new ComparisonResult { Success = true, NewWords = words, InsertedWords = inserted, CrossoverPercent = crossover } }; }
    private static BatchRecord EpDone(BatchPair pair, int descriptionInserted, int claimInserted)
    {
        BatchRecord result = Done(pair, 10, descriptionInserted + claimInserted, 0);
        result.Result.Sections = new[] { Section("Description", 8, descriptionInserted), Section("Full specification", 10, descriptionInserted + claimInserted), Section("Claims", 2, claimInserted) };
        return result;
    }
    private static ComparisonSectionResult Section(string scope, int words, int inserted)
    { return new ComparisonSectionResult { Scope = scope, NewWords = words, InsertedWords = inserted, CrossoverPercent = 100.0 * (words - inserted) / words }; }
    private static void Reject(Action action, string label) { bool threw = false; try { action(); } catch (IOException) { threw = true; } catch (InvalidOperationException) { threw = true; } catch (ArgumentException) { threw = true; } Equal(true, threw, label); }
    private static void Equal(object expected, object actual, string label) { if (!Object.Equals(expected, actual)) throw new Exception(label + ": expected " + expected + "; got " + actual); checks++; }
}
