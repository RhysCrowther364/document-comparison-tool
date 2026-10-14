using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Windows.Forms;
using LocalDocumentCompare;

internal static class BatchUiStateTests
{
    private static int checks;

    [STAThread]
    public static int Main()
    {
        string root = Path.GetFullPath(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "batch-ui-state-" + Guid.NewGuid().ToString("N")));
        Directory.CreateDirectory(root);
        try
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            string first = Path.Combine(root, "first");
            string second = Path.Combine(root, "second");
            Directory.CreateDirectory(first); Directory.CreateDirectory(second);
            File.WriteAllText(Path.Combine(first, "02.pdf"), "synthetic PDF placeholder; never opened");
            File.WriteAllText(Path.Combine(first, "01.docx"), "synthetic Word placeholder; never opened");
            File.WriteAllText(Path.Combine(second, "03.doc"), "synthetic Word placeholder; never opened");
            File.WriteAllText(Path.Combine(second, "01.docx"), "synthetic Word placeholder; never opened");

            using (Form host = new Form { Opacity = 0, ShowInTaskbar = false, Width = 800, Height = 950 })
            using (BatchComparisonControl control = new BatchComparisonControl { Dock = DockStyle.Fill })
            {
                host.Controls.Add(control); host.Show(); Application.DoEvents();
                TextBox folder = Field<TextBox>(control, "folderPath");
                Button start = Field<Button>(control, "startBatch");
                DataGridView grid = Field<DataGridView>(control, "fileGrid");
                List<BatchInput> files = Field<List<BatchInput>>(control, "files");
                Equal(false, start.Enabled, "empty list cannot start");
                folder.Text = "\"" + first + "\"";
                Invoke(control, "LoadFiles", null, EventArgs.Empty);
                Equal(2, files.Count, "quoted folder loads two files");
                Equal("01.docx", Path.GetFileName(files[0].Path), "initial filename ordering");
                Equal(true, start.Enabled, "loaded pair enables start");
                Equal(true, grid.Rows[0].Cells[2].ReadOnly, "Word OCR checkbox is read-only");
                Equal(false, grid.Rows[1].Cells[2].ReadOnly, "PDF OCR checkbox is editable");
                grid.Rows[1].Cells[2].Value = true;
                Equal(true, files[1].UseOcr, "PDF OCR checkbox updates selected input");
                grid.Rows[0].Cells[2].Value = true;
                Equal(false, files[0].UseOcr, "Word cannot acquire OCR programmatically");

                // Seed an already completed batch without comparing documents or starting Word.
                BatchPlan oldPlan = BatchPlan.Create(files, Path.Combine(root, "results"), "WF");
                ComparisonResult oldResult = new ComparisonResult { Success = true, NewWords = 10, InsertedWords = 1, CrossoverPercent = 90 };
                List<BatchRecord> oldRecords = new List<BatchRecord> { new BatchRecord { Pair = oldPlan.Pairs[0], State = "Completed", Result = oldResult } };
                File.WriteAllText(oldPlan.Pairs[0].OutputPath, "synthetic completed comparison placeholder");
                string oldReport = BatchReport.Save(oldPlan, oldRecords);
                string oldReportText = File.ReadAllText(oldReport);
                Set(control, "plan", oldPlan); Set(control, "records", oldRecords); Set(control, "reportPath", oldReport);
                Invoke(control, "RefreshResults");
                DataGridView resultGrid = Field<DataGridView>(control, "resultsGrid");
                resultGrid.Rows[0].Selected = true;
                Invoke(control, "UpdateOutputButtons");
                Equal(true, Field<Button>(control, "openComparison").Enabled, "completed comparison can be opened");
                File.Delete(oldPlan.Pairs[0].OutputPath);
                Invoke(control, "UpdateOutputButtons");
                Equal(true, Field<Button>(control, "openComparison").Enabled, "completed result button uses saved state without checking current file availability");
                File.WriteAllText(oldPlan.Pairs[0].OutputPath, "synthetic completed comparison placeholder");
                Set(control, "reportPath", Path.Combine(root, "temporarily-unavailable-report.csv"));
                Invoke(control, "UpdateOutputButtons");
                Equal(true, Field<Button>(control, "openReport").Enabled, "saved report button uses recorded path without checking current file availability");
                Set(control, "reportPath", oldReport);

                folder.Text = second;
                Equal(0, files.Count, "editing the folder clears stale files");
                Equal(0, grid.Rows.Count, "editing the folder clears the displayed selection");
                Equal(false, start.Enabled, "edited folder requires reload before starting");
                Equal(true, Object.ReferenceEquals(oldPlan, Field<BatchPlan>(control, "plan")), "folder edit retains previous output plan");
                Equal(true, Object.ReferenceEquals(oldRecords, Field<List<BatchRecord>>(control, "records")), "folder edit retains previous results");
                Equal(1, resultGrid.Rows.Count, "folder edit retains completed result rows");
                Equal(true, Field<Button>(control, "openComparison").Enabled, "folder edit keeps completed comparison accessible");
                Equal(true, Field<Button>(control, "openReport").Enabled, "folder edit keeps old report accessible");
                Equal(oldReportText, File.ReadAllText(oldReport), "folder edit does not change the saved report");

                folder.Text = "\"" + first + "\"";
                Invoke(control, "LoadFiles", null, EventArgs.Empty);
                Equal(true, start.Enabled, "reload restores start");
                grid.Rows[1].Cells[2].Value = true;
                grid.ClearSelection(); grid.Rows[1].Selected = true; grid.CurrentCell = grid.Rows[1].Cells[1];
                Invoke(control, "MoveSelected", -1);
                Equal("02.pdf", Path.GetFileName(files[0].Path), "move up changes the actual selected order");
                Equal(true, files[0].UseOcr, "OCR choice moves with its file");
                Invoke(control, "MoveSelected", 1);
                Equal("02.pdf", Path.GetFileName(files[1].Path), "move down restores selected order");
                Field<Button>(control, "removeFile").PerformClick();
                Equal(1, files.Count, "remove deletes the selected input");
                Equal("01.docx", Path.GetFileName(files[0].Path), "remove leaves the other input");
                Equal(false, start.Enabled, "one remaining file cannot start");
                folder.Text = second;
                Invoke(control, "LoadFiles", null, EventArgs.Empty);
                Equal("03.doc", Path.GetFileName(files[1].Path), "reloading a changed folder uses its own files");

                using (FileStream locked = new FileStream(oldReport, FileMode.Open, FileAccess.ReadWrite, FileShare.None))
                {
                    Equal(false, (bool)Invoke(control, "SaveReport"), "locked CSV cannot be saved while idle");
                    Equal(true, Field<TextBox>(control, "details").Text.Contains("Save results table"), "idle save error names its available retry button");
                    Equal(false, Field<Label>(control, "state").Text.Contains("paused"), "idle report error does not claim a paused batch");
                    Invoke(control, "SetBusy", true);
                    int count = files.Count;
                    folder.Text = first;
                    Equal(count, files.Count, "active selection stays frozen while busy");
                    Equal(false, (bool)Invoke(control, "SaveReport"), "locked CSV cannot be saved while busy");
                    Equal(true, Field<TextBox>(control, "details").Text.Contains("Retry saving results"), "busy save error names its available retry button");
                    Equal(true, Field<Label>(control, "state").Text.Contains("paused"), "busy report error identifies paused work");
                    Invoke(control, "SetBusy", false);
                }
                Equal(true, (bool)Invoke(control, "SaveReport"), "CSV save can be retried after unlocking");
                Equal(true, Field<System.Diagnostics.Process>(control, "worker") == null, "all UI state tests ran without a comparison worker");
                host.Close();
            }
            Console.WriteLine("Batch UI state: " + checks + " checks passed; no Word comparisons started.");
            return 0;
        }
        catch (Exception error) { Console.Error.WriteLine(error); return 1; }
        finally
        {
            string parent = Path.GetFullPath(AppDomain.CurrentDomain.BaseDirectory).TrimEnd(Path.DirectorySeparatorChar) + Path.DirectorySeparatorChar;
            if (root.StartsWith(parent, StringComparison.OrdinalIgnoreCase) && Path.GetFileName(root).StartsWith("batch-ui-state-", StringComparison.Ordinal))
                try { Directory.Delete(root, true); } catch { }
        }
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
