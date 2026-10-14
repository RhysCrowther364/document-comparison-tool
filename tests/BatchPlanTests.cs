using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using LocalDocumentCompare;

internal static class BatchPlanTests
{
    private static int checks;
    public static int Main()
    {
        string root = Path.GetFullPath(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "batch-model-" + Guid.NewGuid().ToString("N").Substring(0, 8)));
        Directory.CreateDirectory(root);
        CultureInfo originalCulture = Thread.CurrentThread.CurrentCulture;
        try
        {
            string natural = Path.Combine(root, "natural");
            Directory.CreateDirectory(natural);
            foreach (string name in new[] { "Version10.PDF", "Version2.docx", "Version01.doc", "~$Version3.docx", "notes.txt", "macro.docm" })
                File.WriteAllText(Path.Combine(natural, name), "synthetic");
            Directory.CreateDirectory(Path.Combine(natural, "nested"));
            File.WriteAllText(Path.Combine(natural, "nested", "Version0.docx"), "synthetic");
            List<BatchInput> loaded = BatchPlan.LoadFolder(natural);
            Equal("Version01.doc|Version2.docx|Version10.PDF", String.Join("|", loaded.Select(i => Path.GetFileName(i.Path))), "natural order, supported types and top-level scope");
            Equal(true, loaded.All(i => !i.UseOcr), "OCR defaults to off");

            string inputRoot = Path.Combine(root, "inputs");
            Directory.CreateDirectory(inputRoot);
            var inputs = new List<BatchInput>();
            foreach (string name in new[] { "1.docx", "2.docx", "3.pdf", "4.doc" })
                inputs.Add(Input(inputRoot, name));
            inputs[2].UseOcr = true;
            BatchPlan plan = BatchPlan.Create(inputs, Path.Combine(root, "results"), "wf");
            Equal("WF", plan.JobType, "job type normalized");
            Equal(6, plan.Pairs.Count, "four files produce six pairs");
            Equal("2vs1|3vs1|4vs1|3vs2|4vs2|4vs3",
                String.Join("|", plan.Pairs.Select(p => (p.NewIndex + 1) + "vs" + (p.ArchiveIndex + 1))), "all-pairs archive-first order");
            Equal("2 vs 1.docx", Path.GetFileName(plan.Pairs[0].OutputPath), "exact comparison filename");
            Equal("001", Path.GetFileName(Path.GetDirectoryName(plan.Pairs[0].OutputPath)), "numbered pair folder");
            Equal(true, plan.Pairs.All(p => Directory.Exists(Path.GetDirectoryName(p.OutputPath))), "pair folders created");
            inputs[0].Path = inputs[1].Path;
            inputs[2].UseOcr = false;
            Equal("1.docx", Path.GetFileName(plan.Inputs[0].Path), "plan freezes selected input paths");
            Equal(true, plan.Inputs[2].UseOcr, "plan freezes per-input OCR");
            File.WriteAllText(plan.Pairs[0].OutputPath, "existing result");
            BatchPlan another = BatchPlan.Create(plan.Inputs, Path.Combine(root, "results"), "WF");
            Equal(false, plan.OutputDirectory == another.OutputDirectory, "new batch has a unique result folder");
            Equal("existing result", File.ReadAllText(plan.Pairs[0].OutputPath), "existing comparison never overwritten by planning");

            var sameStems = new[] { Input(inputRoot, "same.docx"), Input(inputRoot, "same.doc"), Input(inputRoot, "same.pdf") };
            BatchPlan same = BatchPlan.Create(sameStems, Path.Combine(root, "results"), "WF");
            Equal(3, same.Pairs.Select(p => p.OutputPath).Distinct(StringComparer.OrdinalIgnoreCase).Count(), "same stem inputs have separate output paths");
            Equal(true, same.Pairs.All(p => Path.GetFileName(p.OutputPath) == "same vs same.docx"), "same stem filenames remain exact");

            string invalid = Path.Combine(root, "must-not-exist");
            Reject(delegate { BatchPlan.Create(new[] { plan.Inputs[0], new BatchInput { Path = Path.Combine(inputRoot, ".", "1.docx") } }, invalid, "WF"); }, "normalized duplicate rejected");
            Reject(delegate { BatchPlan.Create(new[] { plan.Inputs[0] }, invalid, "WF"); }, "one input rejected");
            Reject(delegate { BatchPlan.Create(plan.Inputs, invalid, "invalid"); }, "unknown job type rejected");
            Reject(delegate { BatchPlan.Create(new[] { plan.Inputs[0], new BatchInput { Path = plan.Inputs[1].Path, UseOcr = true } }, invalid, "WF"); }, "OCR on Word rejected");
            Reject(delegate { BatchPlan.Create(new[] { plan.Inputs[0], new BatchInput { Path = Path.Combine(inputRoot, "missing.pdf") } }, invalid, "WF"); }, "missing file rejected");
            Reject(delegate { BatchPlan.Create(new[] { plan.Inputs[0], Input(inputRoot, "unsupported.txt") }, invalid, "WF"); }, "unsupported file rejected");
            Reject(delegate { BatchPlan.Create(new[] { plan.Inputs[0], Input(inputRoot, "~$lock.docx") }, invalid, "WF"); }, "Word lock file rejected");
            string longStem = new String('a', 90);
            Reject(delegate { BatchPlan.Create(new[] { Input(inputRoot, longStem + ".doc"), Input(inputRoot, longStem + ".pdf") }, invalid, "WF"); }, "long output path rejected before creation");
            Equal(false, Directory.Exists(invalid), "validation failures create no output directories");

            string pendingReport = BatchReport.Save(plan, new List<BatchRecord>());
            List<string[]> pendingRows = Csv(pendingReport);
            Equal(7, pendingRows.Count, "report includes every pair even without explicit records");
            Equal(true, pendingRows.Skip(1).All(r => r[4] == "Pending" && r[5] == "" && r[6] == "" && r[7] == ""), "pending figures are blank");
            byte[] bytes = File.ReadAllBytes(pendingReport);
            Equal(true, bytes.Length > 3 && bytes[0] == 0xef && bytes[1] == 0xbb && bytes[2] == 0xbf, "Excel UTF8 BOM");
            Equal("Pair|Archive|New|Scope|Status|Crossover (%)|New file word count|Highlighted insertion word count|Comparison document|Notes", String.Join("|", pendingRows[0]), "exact report columns");

            var records = new List<BatchRecord> {
                new BatchRecord { Pair = plan.Pairs[0], State = "Completed", Result = Success(400, 5, 98.75) },
                new BatchRecord { Pair = plan.Pairs[1], State = "Completed", Result = new ComparisonResult { Success = false, LanguageReviewRequired = true, Error = "Manual language review", NewWords = 999, InsertedWords = 999, CrossoverPercent = 12 } },
                new BatchRecord { Pair = plan.Pairs[2], State = "Failed", Result = new ComparisonResult { Error = "Conversion failed", NewWords = 999, CrossoverPercent = 12 } },
                new BatchRecord { Pair = plan.Pairs[3], State = "Cancelled", Result = Success(999, 999, 12) },
                new BatchRecord { Pair = plan.Pairs[4], State = "Running" },
                new BatchRecord { Pair = plan.Pairs[5], State = "Not run" }
            };
            Thread.CurrentThread.CurrentCulture = CultureInfo.GetCultureInfo("fr-FR");
            BatchReport.Save(plan, records);
            List<string[]> rows = Csv(plan.ReportPath);
            Equal("98.8", rows[1][5], "crossover uses one invariant decimal");
            Equal("400|5", rows[1][6] + "|" + rows[1][7], "completed counts emitted");
            Equal(Path.Combine("001", "2 vs 1.docx"), rows[1][8], "comparison path is relative to batch folder");
            Equal("Review required", rows[2][4], "manual review cannot be reported as completion");
            Equal(true, rows.Skip(2).All(r => r[5] == "" && r[6] == "" && r[7] == "" && r[8] == ""), "all noncompleted metrics and document paths blank");
            Equal("Conversion failed", rows[3][9], "actual error included");
            Equal(0, Directory.GetFiles(plan.OutputDirectory, "*.tmp").Length, "atomic saves leave no temporary CSV");
            records[1].State = "Failed";
            BatchReport.Save(plan, records);
            string[] skippedReview = Csv(plan.ReportPath)[2];
            Equal("Failed", skippedReview[4], "explicitly skipped language review remains failed");
            Equal("Manual language review", skippedReview[9], "skipped review retains its original explanation");
            Equal("", skippedReview[6], "skipped review never exports partial metrics");
            string prior = File.ReadAllText(plan.ReportPath);
            Reject(delegate { BatchReport.Save(plan, new[] { records[0], records[0] }); }, "duplicate report record rejected");
            Equal(prior, File.ReadAllText(plan.ReportPath), "report validation failure preserves previous CSV");

            BatchPlan formula = BatchPlan.Create(new[] { Input(inputRoot, "=Archive.docx"), Input(inputRoot, "+New,version.docx") }, Path.Combine(root, "results"), "WF");
            ComparisonResult formulaResult = Success(10, 2, 80);
            formulaResult.Warnings = new[] { " \t=SUM(1,2)\r\n\"quoted\", extra" };
            BatchReport.Save(formula, new[] { new BatchRecord { Pair = formula.Pairs[0], State = "Completed", Result = formulaResult } });
            string[] formulaRow = Csv(formula.ReportPath)[1];
            Equal("'=Archive.docx", formulaRow[1], "formula-leading archive filename made literal");
            Equal("'+New,version.docx", formulaRow[2], "formula-leading new filename and comma escaped");
            Equal("' \t=SUM(1,2)\r\n\"quoted\", extra", formulaRow[9], "formula-safe notes preserve quotes and multiline content");

            BatchPlan ep = BatchPlan.Create(new[] { plan.Inputs[0], plan.Inputs[1] }, Path.Combine(root, "results"), "EP");
            ComparisonResult epResult = Success(30, 3, 90);
            epResult.Sections = new[] {
                new ComparisonSectionResult { Scope = "Claims", NewWords = 10, InsertedWords = 1, CrossoverPercent = 90 },
                new ComparisonSectionResult { Scope = "Description", NewWords = 20, InsertedWords = 2, CrossoverPercent = 90 },
                new ComparisonSectionResult { Scope = "Full specification", NewWords = 30, InsertedWords = 3, CrossoverPercent = 90 }
            };
            var epRecord = new BatchRecord { Pair = ep.Pairs[0], State = "Completed", Result = epResult };
            BatchReport.Save(ep, new[] { epRecord });
            List<string[]> epRows = Csv(ep.ReportPath);
            Equal(4, epRows.Count, "completed EP pair has three rows");
            Equal("Description|Full specification|Claims", String.Join("|", epRows.Skip(1).Select(r => r[3])), "EP report scope order is stable");
            Equal("20|30|10", String.Join("|", epRows.Skip(1).Select(r => r[6])), "EP rows contain correct scope figures");
            epRecord.State = "Failed";
            epResult.Error = "EP conversion failed";
            BatchReport.Save(ep, new[] { epRecord });
            epRows = Csv(ep.ReportPath);
            Equal(2, epRows.Count, "failed EP pair has one status row");
            Equal("", epRows[1][6], "failed EP figures are blank even with prior metrics");

            BatchPlan protectedReport = BatchPlan.Create(new[] { plan.Inputs[0], plan.Inputs[1] }, Path.Combine(root, "results"), "WF");
            File.WriteAllText(protectedReport.ReportPath, "unrelated existing file");
            Reject(delegate { BatchReport.Save(protectedReport, new List<BatchRecord>()); }, "unowned report not overwritten");
            Equal("unrelated existing file", File.ReadAllText(protectedReport.ReportPath), "unowned report contents preserved");
            Console.WriteLine("PASS: " + checks + " batch planning/reporting checks. Fixtures: " + root);
            return 0;
        }
        catch (Exception ex) { Console.Error.WriteLine("FAIL: " + ex); return 1; }
        finally { Thread.CurrentThread.CurrentCulture = originalCulture; }
    }

    private static BatchInput Input(string directory, string name)
    {
        string path = Path.Combine(directory, name);
        if (!File.Exists(path)) File.WriteAllText(path, "synthetic");
        return new BatchInput { Path = path };
    }
    private static ComparisonResult Success(int words, int inserted, double crossover)
    {
        return new ComparisonResult { Success = true, NewWords = words, InsertedWords = inserted, CrossoverPercent = crossover, Warnings = new string[0] };
    }
    private static void Equal<T>(T expected, T actual, string label)
    {
        if (!Object.Equals(expected, actual)) throw new Exception(label + ": expected " + expected + "; got " + actual);
        checks++;
    }
    private static void Reject(Action action, string label)
    {
        try { action(); }
        catch (Exception) { checks++; return; }
        throw new Exception(label + ": unexpectedly accepted");
    }
    private static List<string[]> Csv(string path)
    {
        string text = File.ReadAllText(path, Encoding.UTF8);
        var rows = new List<string[]>();
        var fields = new List<string>();
        var field = new StringBuilder();
        bool quoted = false;
        for (int i = 0; i < text.Length; i++)
        {
            char value = text[i];
            if (value == '"')
            {
                if (quoted && i + 1 < text.Length && text[i + 1] == '"') { field.Append('"'); i++; }
                else quoted = !quoted;
            }
            else if (!quoted && value == ',') { fields.Add(field.ToString()); field.Length = 0; }
            else if (!quoted && (value == '\r' || value == '\n'))
            {
                if (value == '\r' && i + 1 < text.Length && text[i + 1] == '\n') i++;
                fields.Add(field.ToString()); rows.Add(fields.ToArray()); fields.Clear(); field.Length = 0;
            }
            else field.Append(value);
        }
        if (quoted) throw new Exception("Malformed quoted CSV");
        if (field.Length > 0 || fields.Count > 0) { fields.Add(field.ToString()); rows.Add(fields.ToArray()); }
        if (rows.Any(row => row.Length != 10)) throw new Exception("CSV column count is not ten");
        return rows;
    }
}
