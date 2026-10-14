using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Security.Cryptography;
using System.Windows.Forms;
using LocalDocumentCompare;

internal static class BatchResetUiTests
{
    private static int checks;

    [STAThread]
    public static int Main()
    {
        string root = Path.GetFullPath(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "batch-reset-ui-" + Guid.NewGuid().ToString("N")));
        Directory.CreateDirectory(root);
        try
        {
            Application.EnableVisualStyles(); Application.SetCompatibleTextRenderingDefault(false);
            string input = Path.Combine(root, "inputs"); Directory.CreateDirectory(input);
            string archive = Path.Combine(input, "01.docx"), revised = Path.Combine(input, "02.pdf");
            File.WriteAllText(archive, "Synthetic Word placeholder; never opened");
            File.WriteAllText(revised, "Synthetic PDF placeholder; never opened");
            string archiveHash = Hash(archive), revisedHash = Hash(revised);
            using (Form host = new Form { Opacity = 0, ShowInTaskbar = false, Width = 760, Height = 950 })
            using (BatchComparisonControl control = new BatchComparisonControl { Dock = DockStyle.Fill })
            {
                host.Controls.Add(control); host.Show(); Application.DoEvents();
                Button reset = Field<Button>(control, "resetBatch");
                ComboBox jobType = Field<ComboBox>(control, "jobType");
                TextBox output = Field<TextBox>(control, "outputPath");
                Panel configuration = Field<Panel>(control, "configuration");
                jobType.SelectedIndex = 1;
                const string destination = @"\\server\team\Pilot results";
                output.Text = destination; // UI preference only; tests write nothing to this path.
                Equal(true, reset.Visible && reset.Enabled, "Reset is visible and enabled before selecting files");
                Equal(true, reset.Right <= configuration.ClientSize.Width && jobType.Right + 16 <= reset.Left, "Reset and job type fit the minimum-width host");
                host.Width = 1040; Application.DoEvents();
                Equal(true, reset.Right <= configuration.ClientSize.Width && jobType.Right + 16 <= reset.Left, "Reset and job type stay separated after widening");
                host.Width = 760; Application.DoEvents();

                Load(control, input);
                Field<DataGridView>(control, "fileGrid").Rows[1].Cells[2].Value = true;
                Equal(true, Field<List<BatchInput>>(control, "files")[1].UseOcr, "idle test has a selected PDF OCR option");
                int generation = Field<int>(control, "runGeneration");
                reset.PerformClick();
                AssertReset(control, generation, destination, "idle selection");
                Load(control, input);
                Equal(false, Field<List<BatchInput>>(control, "files")[1].UseOcr, "reloading after Reset starts with no stale PDF OCR selection");

                BatchPlan saved = BatchPlan.Create(Field<List<BatchInput>>(control, "files"), Path.Combine(root, "saved-results"), "EP");
                string comparison = saved.Pairs[0].OutputPath, summary = Path.ChangeExtension(comparison, ".txt");
                File.WriteAllText(comparison, "Saved synthetic comparison; retain on Reset");
                File.WriteAllText(summary, "Saved synthetic summary; retain on Reset");
                var completed = new ComparisonResult { Success = true, NewWords = 12, InsertedWords = 3, CrossoverPercent = 75 };
                var savedRecords = new List<BatchRecord> { new BatchRecord { Pair = saved.Pairs[0], State = "Completed", Result = completed } };
                string report = BatchReport.Save(saved, savedRecords);
                string comparisonHash = Hash(comparison), summaryHash = Hash(summary), reportHash = Hash(report);
                foreach (string state in new[] { "Completed", "Cancelled", "Failed" })
                {
                    Load(control, input);
                    SeedResult(control, saved, report, state, completed);
                    Set(control, "pendingStart", true); Set(control, "pendingStartConfirmed", true);
                    Set(control, "nextScheduled", true); Set(control, "cancelRequested", true);
                    Set(control, "reviewPaused", true); Set(control, "reportBlocked", true); Set(control, "stopBlocked", true);
                    Field<Label>(control, "state").Text = "Previous " + state;
                    Field<TextBox>(control, "details").Text = "Previous error/review/output information";
                    Field<System.Windows.Forms.Timer>(control, "timer").Interval = 600000;
                    Field<System.Windows.Forms.Timer>(control, "timer").Start();
                    generation = Field<int>(control, "runGeneration");
                    reset.PerformClick();
                    AssertReset(control, generation, destination, state);
                    Equal(comparisonHash, Hash(comparison), state + " Reset preserves the saved comparison");
                    Equal(summaryHash, Hash(summary), state + " Reset preserves the saved summary");
                    Equal(reportHash, Hash(report), state + " Reset preserves the saved CSV");
                }

                // An ordinary cleanup failure must remain visible, without touching its files.
                Load(control, input);
                string leftover = Path.Combine(root, "retained-temp"); Directory.CreateDirectory(leftover);
                string leftoverFile = Path.Combine(leftover, "leftover.txt"); File.WriteAllText(leftoverFile, "Retained temporary placeholder");
                Set(control, "jobDirectory", leftover); Set(control, "keepJobDirectory", false);
                generation = Field<int>(control, "runGeneration");
                reset.PerformClick();
                AssertReset(control, generation, destination, "cleanup warning");
                Equal(leftover, Field<string>(control, "jobDirectory"), "Reset keeps the retained temporary path");
                Equal(true, File.Exists(leftoverFile), "Reset does not delete retained temporary files");
                Equal(true, Field<TextBox>(control, "details").Text.Contains("Previous temporary files remain: " + leftover), "Reset retains the temporary cleanup note");
                Set(control, "jobDirectory", null);

                Load(control, input); SeedResult(control, saved, report, "Completed", completed);
                Invoke(control, "SetBusy", true);
                Equal(false, reset.Enabled, "Reset is disabled during an active batch");
                AssertGuard(control, "running batch");
                Set(control, "reviewPaused", true); Invoke(control, "UpdateActions");
                Equal(false, reset.Enabled, "Reset remains disabled during manual language review");
                AssertGuard(control, "language review");
                Set(control, "reviewPaused", false); Set(control, "reportBlocked", true); Invoke(control, "UpdateActions");
                Equal(false, reset.Enabled, "Reset remains disabled during report retry pause");
                AssertGuard(control, "report saving pause");
                Set(control, "reportBlocked", false); Set(control, "stopBlocked", true); Invoke(control, "UpdateActions");
                Equal(false, reset.Enabled, "Reset remains disabled during stop retry pause");
                AssertGuard(control, "stop pause");
                Invoke(control, "SetBusy", false);
                Set(control, "stopBlocked", false);

                using (Process placeholder = new Process())
                {
                    Set(control, "worker", placeholder); Invoke(control, "UpdateActions");
                    Equal(false, reset.Enabled, "Reset is disabled while a worker reference exists");
                    AssertGuard(control, "worker reference");
                    Set(control, "worker", null);
                }
                Set(control, "disposing", true); Invoke(control, "UpdateActions");
                Equal(false, reset.Enabled, "Reset is disabled during disposal");
                AssertGuard(control, "disposal");
                Set(control, "disposing", false);
                Set(control, "jobDirectory", leftover); Set(control, "keepJobDirectory", true); Invoke(control, "UpdateActions");
                Equal(false, reset.Enabled, "Reset is disabled when an owned child has not been confirmed stopped");
                AssertGuard(control, "unconfirmed child stop");
                Equal(leftover, Field<string>(control, "jobDirectory"), "blocked Reset retains the owned temporary path");
                Equal(true, Field<bool>(control, "keepJobDirectory"), "blocked Reset retains stop protection");
                Set(control, "jobDirectory", null); Set(control, "keepJobDirectory", false); Invoke(control, "UpdateActions");

                // Queue an old next-pair callback, reset, then simulate a new busy generation.
                Set(control, "nextScheduled", false); Set(control, "cancelRequested", false);
                Set(control, "reviewPaused", false); Set(control, "reportBlocked", false); Set(control, "stopBlocked", false);
                Invoke(control, "SetBusy", true);
                Invoke(control, "ScheduleNext");
                Equal(true, Field<bool>(control, "nextScheduled"), "the stale callback test queued a next pair");
                Invoke(control, "SetBusy", false);
                generation = Field<int>(control, "runGeneration");
                reset.PerformClick();
                AssertReset(control, generation, destination, "queued callback");
                Invoke(control, "SetBusy", true); Application.DoEvents();
                Equal(true, control.IsBusy, "an old queued callback does not complete a newer batch generation");
                Equal(-1, Field<int>(control, "currentIndex"), "an old queued callback does not advance the new pair index");
                Equal(true, Field<Process>(control, "worker") == null, "Reset tests never start a comparison worker");
                Invoke(control, "SetBusy", false);
                Equal(archiveHash, Hash(archive), "all Reset operations preserve archive bytes");
                Equal(revisedHash, Hash(revised), "all Reset operations preserve new-file bytes");
                host.Close();
            }
            Console.WriteLine("Batch Reset UI: " + checks + " checks passed; no workers or Word comparisons started.");
            return 0;
        }
        catch (Exception error) { Console.Error.WriteLine(error); return 1; }
        finally
        {
            string parent = Path.GetFullPath(AppDomain.CurrentDomain.BaseDirectory).TrimEnd(Path.DirectorySeparatorChar) + Path.DirectorySeparatorChar;
            if (root.StartsWith(parent, StringComparison.OrdinalIgnoreCase) && Path.GetFileName(root).StartsWith("batch-reset-ui-", StringComparison.Ordinal))
                try { Directory.Delete(root, true); } catch { }
        }
    }

    private static void Load(BatchComparisonControl control, string input)
    {
        Field<TextBox>(control, "folderPath").Text = input;
        Invoke(control, "LoadFiles", null, EventArgs.Empty);
        if (Field<List<BatchInput>>(control, "files").Count != 2) throw new Exception("Synthetic input folder did not load two files.");
    }

    private static void SeedResult(BatchComparisonControl control, BatchPlan plan, string report, string state, ComparisonResult completed)
    {
        Set(control, "plan", plan); Set(control, "reportPath", report); Set(control, "currentIndex", 0);
        Set(control, "records", new List<BatchRecord> { new BatchRecord { Pair = plan.Pairs[0], State = state,
            Result = state == "Completed" ? completed : new ComparisonResult { Success = false, Error = "Synthetic " + state } } });
        Invoke(control, "RefreshResults");
        Field<DataGridView>(control, "resultsGrid").Rows[0].Selected = true;
        Invoke(control, "UpdateActions");
    }

    private static void AssertReset(BatchComparisonControl control, int previousGeneration, string destination, string scenario)
    {
        Equal("", Field<TextBox>(control, "folderPath").Text, scenario + " clears the folder selection");
        Equal(0, Field<List<BatchInput>>(control, "files").Count, scenario + " clears files and their OCR choices");
        Equal(0, Field<DataGridView>(control, "fileGrid").Rows.Count, scenario + " clears displayed inputs");
        Equal(0, Field<DataGridView>(control, "resultsGrid").Rows.Count, scenario + " clears displayed results");
        Equal(0, Field<List<BatchRecord>>(control, "records").Count, scenario + " clears previous result records");
        Equal(true, Field<BatchPlan>(control, "plan") == null && Field<string>(control, "reportPath") == null, scenario + " clears previous output references");
        Equal(1, Field<ComboBox>(control, "jobType").SelectedIndex, scenario + " preserves EP mode");
        Equal(destination, Field<TextBox>(control, "outputPath").Text, scenario + " preserves the output location");
        Equal(previousGeneration + 1, Field<int>(control, "runGeneration"), scenario + " invalidates queued callbacks");
        Equal(-1, Field<int>(control, "currentIndex"), scenario + " clears the pair index");
        Equal(false, Field<System.Windows.Forms.Timer>(control, "timer").Enabled, scenario + " stops polling");
        Equal(false, Field<bool>(control, "nextScheduled") || Field<bool>(control, "pendingStart") || Field<bool>(control, "pendingStartConfirmed") || Field<bool>(control, "reviewPaused") || Field<bool>(control, "reportBlocked") || Field<bool>(control, "stopBlocked") || Field<bool>(control, "cancelRequested"), scenario + " clears previous session flags");
        Equal(false, Field<Button>(control, "startBatch").Enabled || Field<Button>(control, "openComparison").Enabled || Field<Button>(control, "openFolder").Enabled || Field<Button>(control, "openReport").Enabled || Field<Button>(control, "saveResults").Enabled, scenario + " disables actions for old outputs or missing inputs");
        Equal(false, Field<Button>(control, "languageChecked").Visible || Field<Button>(control, "skipPair").Visible || Field<Button>(control, "retryReport").Visible, scenario + " hides previous review actions");
        Equal(true, Field<Label>(control, "state").Text.StartsWith("Ready"), scenario + " restores Ready state");
        Equal(true, Field<TextBox>(control, "folderPath").Focused, scenario + " focuses the folder input");
    }

    private static void AssertGuard(BatchComparisonControl control, string scenario)
    {
        int count = Field<List<BatchInput>>(control, "files").Count, generation = Field<int>(control, "runGeneration");
        string folder = Field<TextBox>(control, "folderPath").Text;
        BatchPlan plan = Field<BatchPlan>(control, "plan");
        Invoke(control, "ResetBatch");
        Equal(count, Field<List<BatchInput>>(control, "files").Count, scenario + " leaves selected inputs unchanged");
        Equal(folder, Field<TextBox>(control, "folderPath").Text, scenario + " leaves the selected folder unchanged");
        Equal(generation, Field<int>(control, "runGeneration"), scenario + " leaves generation unchanged");
        Equal(true, Object.ReferenceEquals(plan, Field<BatchPlan>(control, "plan")), scenario + " retains the active output plan");
    }

    private static string Hash(string path)
    {
        using (SHA256 hash = SHA256.Create()) using (Stream stream = File.OpenRead(path)) return Convert.ToBase64String(hash.ComputeHash(stream));
    }
    private static T Field<T>(object instance, string name) { return (T)instance.GetType().GetField(name, BindingFlags.Instance | BindingFlags.NonPublic).GetValue(instance); }
    private static void Set(object instance, string name, object value) { instance.GetType().GetField(name, BindingFlags.Instance | BindingFlags.NonPublic).SetValue(instance, value); }
    private static object Invoke(object instance, string name, params object[] arguments) { return instance.GetType().GetMethod(name, BindingFlags.Instance | BindingFlags.NonPublic).Invoke(instance, arguments); }
    private static void Equal<T>(T expected, T actual, string name)
    {
        if (!EqualityComparer<T>.Default.Equals(expected, actual)) throw new Exception(name + ": expected " + expected + ", got " + actual);
        checks++;
    }
}
