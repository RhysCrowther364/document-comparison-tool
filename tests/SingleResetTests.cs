using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Security.Cryptography;
using System.Windows.Forms;
using LocalDocumentCompare;

internal static class SingleResetTests
{
    private static int checks;

    [STAThread]
    public static int Main()
    {
        string testParent = Path.GetFullPath(AppDomain.CurrentDomain.BaseDirectory).TrimEnd(Path.DirectorySeparatorChar) + Path.DirectorySeparatorChar;
        string root = Path.Combine(testParent, "single-reset-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        try
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            string savedComparison = Path.Combine(root, "saved comparison.docx");
            string savedSummary = Path.ChangeExtension(savedComparison, ".txt");
            File.WriteAllText(savedComparison, "Synthetic saved comparison placeholder; never opened in Word.");
            File.WriteAllText(savedSummary, "Synthetic saved summary: 80.0%, 100 words, 20 inserted words.");
            string comparisonHash = Hash(savedComparison);
            string summaryHash = Hash(savedSummary);

            using (MainForm form = new MainForm())
            {
                form.Opacity = 0; form.ShowInTaskbar = false;
                form.Show(); Application.DoEvents();
                TextBox archive = Field<TextBox>(form, "archive");
                TextBox revised = Field<TextBox>(form, "revised");
                ComboBox jobType = Field<ComboBox>(form, "jobType");
                TextBox destination = Field<TextBox>(form, "destination");
                Panel inputs = Field<Panel>(form, "inputs");
                Button reset = Field<Button>(form, "reset");
                destination.Text = root;
                Equal("Reset", reset.Text, "reset is discoverable by its label");
                Equal("resetSingle", reset.Name, "single reset has a stable named control");
                Equal(true, reset.Visible && reset.Enabled, "idle reset is visible and available");
                Equal(true, reset.AccessibleDescription.Contains("saved documents"), "accessible description explains saved documents are retained");

                // Seed a completed full-document display, then exercise the real click handler.
                SeedInputs(form, root);
                Invoke(form, "ShowResultMetrics", new ComparisonResult { Success = true, NewWords = 100, InsertedWords = 20, CrossoverPercent = 80 });
                Equal(true, Field<Label>(form, "metrics").Text.Contains("80.0%"), "completed full-document statistics are visible before resetting");
                SeedOldState(form, savedComparison);
                reset.PerformClick();
                CheckClean(form, 0, root, "completed full-document reset");
                Equal(comparisonHash, Hash(savedComparison), "full-document reset preserves saved comparison bytes");
                Equal(summaryHash, Hash(savedSummary), "full-document reset preserves saved summary bytes");

                // Section results must also disappear while the selected EP mode remains.
                jobType.SelectedIndex = 1;
                SeedInputs(form, root);
                Invoke(form, "ShowResultMetrics", new ComparisonResult {
                    Success = true, Sections = new[] {
                        new ComparisonSectionResult { Scope = "Description", NewWords = 80, InsertedWords = 10, CrossoverPercent = 87.5 },
                        new ComparisonSectionResult { Scope = "Full specification", NewWords = 100, InsertedWords = 20, CrossoverPercent = 80 },
                        new ComparisonSectionResult { Scope = "Claims", NewWords = 20, InsertedWords = 10, CrossoverPercent = 50 }
                    }
                });
                Equal(3, Field<DataGridView>(form, "sectionMetrics").Rows.Count, "EP results are populated before resetting");
                Equal(true, Field<DataGridView>(form, "sectionMetrics").Visible, "EP results table is visible before resetting");
                SeedOldState(form, savedComparison);
                reset.PerformClick();
                CheckClean(form, 1, root, "completed EP reset");
                Equal(comparisonHash, Hash(savedComparison), "EP reset preserves saved comparison bytes");
                Equal(summaryHash, Hash(savedSummary), "EP reset preserves saved summary bytes");
                Equal(true, Field<Label>(form, "jobHint").Text.Contains("EP needs"), "EP guidance follows the preserved mode");

                // Resetting a stopped language review removes consent and errors for the next job.
                SeedInputs(form, root);
                Field<Label>(form, "state").Text = "Language check required — comparison stopped.";
                Field<TextBox>(form, "details").Text = "Synthetic error and manual-review instructions.";
                Field<Button>(form, "languageChecked").Visible = true;
                Equal(true, Field<Button>(form, "languageChecked").Visible, "manual review action exists before reset");
                reset.PerformClick();
                CheckClean(form, 1, root, "manual-review reset");
                Equal(false, Field<TextBox>(form, "details").Text.Contains("Synthetic error"), "old error text is removed");

                // All of the file controls work again on the same form after clearing.
                archive.Text = Path.Combine(root, "next archive.pdf");
                revised.Text = Path.Combine(root, "next new.pdf");
                Equal(true, Field<CheckBox>(form, "archiveOcr").Enabled && Field<CheckBox>(form, "newOcr").Enabled, "new PDF selections re-enable OCR controls");
                Field<CheckBox>(form, "archiveOcr").Checked = true;
                Field<CheckBox>(form, "newOcr").Checked = true;
                Equal(true, Named<Button>(inputs, "browseArchive").Enabled && Named<Button>(inputs, "browseNew").Enabled, "input browsers remain available after reset");
                Equal(true, Named<Button>(inputs, "browseOutput").Enabled, "destination browser remains available after reset");
                Equal(true, archive.AllowDrop && revised.AllowDrop, "drag and drop remains available after reset");
                Equal(true, Field<Button>(form, "compare").Enabled, "comparison action is ready for the new selections");
                revised.Text = Path.Combine(root, "next new.docx");
                Equal(false, Field<CheckBox>(form, "newOcr").Enabled || Field<CheckBox>(form, "newOcr").Checked, "switching new file to Word clears PDF OCR normally");

                // A Process object need not be launched to prove the active-worker guard.
                using (Process notStarted = new Process())
                {
                    try
                    {
                        Set(form, "worker", notStarted);
                        Invoke(form, "UpdateModeAvailability");
                        Equal(false, reset.Enabled, "active single-worker reference disables Reset");
                        Invoke(form, "ResetComparison");
                        Equal(Path.Combine(root, "next archive.pdf"), archive.Text, "direct reset is guarded during single comparison");
                        Equal(Path.Combine(root, "next new.docx"), revised.Text, "guard keeps the running new-file selection");
                        Equal(true, Object.ReferenceEquals(notStarted, Field<Process>(form, "worker")), "guard retains worker ownership without starting or stopping a process");
                    }
                    finally { Set(form, "worker", null); Invoke(form, "UpdateModeAvailability"); }
                }

                BatchComparisonControl batch = Field<BatchComparisonControl>(form, "batch");
                try
                {
                    Invoke(batch, "SetBusy", true);
                    Equal(false, reset.Enabled, "another active batch disables single Reset");
                    Invoke(form, "ResetComparison");
                    Equal(Path.Combine(root, "next archive.pdf"), archive.Text, "direct single reset is guarded while the batch is busy");
                }
                finally { Invoke(batch, "SetBusy", false); }
                Equal(true, reset.Enabled, "single Reset becomes available when the batch finishes");

                try
                {
                    Set(form, "keepJobDirectory", true);
                    Invoke(form, "UpdateModeAvailability");
                    Equal(false, reset.Enabled, "a cleanup ownership block disables Reset");
                    Invoke(form, "ResetComparison");
                    Equal(Path.Combine(root, "next archive.pdf"), archive.Text, "direct reset preserves state while cleanup is blocked");
                }
                finally { Set(form, "keepJobDirectory", false); Invoke(form, "UpdateModeAvailability"); }

                // Unremoved temporary files are retained, with their location carried into the fresh UI.
                string leftover = Path.Combine(root, "leftover-job");
                Directory.CreateDirectory(leftover);
                string tempDocument = Path.Combine(leftover, "source-copy.docx");
                File.WriteAllText(tempDocument, "Synthetic temporary copy that Reset must not delete.");
                string tempHash = Hash(tempDocument);
                Set(form, "jobDirectory", leftover);
                reset.PerformClick();
                Equal(leftover, Field<string>(form, "jobDirectory"), "reset retains leftover temporary-job ownership reference");
                Equal(tempHash, Hash(tempDocument), "reset does not delete or change leftover temporary files");
                Equal(true, Field<TextBox>(form, "details").Text.Contains(leftover), "fresh screen retains the leftover temporary-folder note");
                Equal(comparisonHash, Hash(savedComparison), "repeated resets preserve saved comparison bytes");
                Equal(summaryHash, Hash(savedSummary), "repeated resets preserve saved summary bytes");

                form.Size = form.MinimumSize; Application.DoEvents();
                InBounds(reset, inputs, "Reset fits at the minimum window size");
                InBounds(jobType, inputs, "job-type dropdown fits at the minimum window size");
                Equal(false, reset.Bounds.IntersectsWith(jobType.Bounds), "Reset does not overlap the job-type dropdown");
                Equal(true, reset.Left - jobType.Right >= 8, "Reset and job-type dropdown retain a readable gap");
                Equal(true, Field<Label>(form, "jobHint").Top >= reset.Bottom, "Reset does not overlap the guidance row");
                Equal(true, Field<Process>(form, "worker") == null && Field<Process>(batch, "worker") == null, "no real single or batch worker was started");
                Set(form, "jobDirectory", null);
                form.Close();
            }
            Console.WriteLine("Single Reset: " + checks + " checks passed; saved bytes preserved; no Word or network use.");
            return 0;
        }
        catch (Exception error) { Console.Error.WriteLine(error); return 1; }
        finally
        {
            string resolved = Path.GetFullPath(root);
            if (resolved.StartsWith(testParent, StringComparison.OrdinalIgnoreCase) && Path.GetFileName(resolved).StartsWith("single-reset-", StringComparison.Ordinal))
                try { Directory.Delete(resolved, true); } catch { }
        }
    }

    private static void SeedInputs(MainForm form, string root)
    {
        Field<TextBox>(form, "archive").Text = Path.Combine(root, "old archive.pdf");
        Field<TextBox>(form, "revised").Text = Path.Combine(root, "old new.pdf");
        Field<CheckBox>(form, "archiveOcr").Checked = true;
        Field<CheckBox>(form, "newOcr").Checked = true;
    }

    private static void SeedOldState(MainForm form, string savedComparison)
    {
        Set(form, "lastOutput", savedComparison);
        Set(form, "cancelling", true);
        Set(form, "started", DateTime.UtcNow);
        Field<Button>(form, "openDocument").Enabled = true;
        Field<Button>(form, "openFolder").Enabled = true;
        Field<Button>(form, "cancel").Enabled = true;
        Field<Label>(form, "state").Text = "Old comparison status.";
        Field<TextBox>(form, "details").Text = "Old comparison details.";
        ProgressBar progress = Field<ProgressBar>(form, "progress");
        progress.Visible = true; progress.Style = ProgressBarStyle.Marquee; progress.Value = 41;
        Field<System.Windows.Forms.Timer>(form, "timer").Start();
    }

    private static void CheckClean(MainForm form, int mode, string destination, string name)
    {
        Equal(String.Empty, Field<TextBox>(form, "archive").Text, name + ": archive cleared");
        Equal(String.Empty, Field<TextBox>(form, "revised").Text, name + ": new file cleared");
        Equal(false, Field<CheckBox>(form, "archiveOcr").Checked || Field<CheckBox>(form, "newOcr").Checked, name + ": OCR selection cleared");
        Equal(false, Field<CheckBox>(form, "archiveOcr").Enabled || Field<CheckBox>(form, "newOcr").Enabled, name + ": OCR unavailable without PDFs");
        Equal(mode, Field<ComboBox>(form, "jobType").SelectedIndex, name + ": selected mode preserved");
        Equal(destination, Field<TextBox>(form, "destination").Text, name + ": save location preserved");
        Equal(false, Field<Button>(form, "languageChecked").Visible, name + ": manual language action hidden");
        Equal(null, Field<string>(form, "lastOutput"), name + ": old output reference cleared");
        Equal(false, Field<Button>(form, "openDocument").Enabled || Field<Button>(form, "openFolder").Enabled, name + ": old output buttons disabled");
        Equal(false, Field<bool>(form, "cancelling"), name + ": cancellation state cleared");
        Equal(default(DateTime), Field<DateTime>(form, "started"), name + ": old start timestamp cleared");
        Equal(false, Field<System.Windows.Forms.Timer>(form, "timer").Enabled, name + ": polling stopped");
        ProgressBar progress = Field<ProgressBar>(form, "progress");
        Equal(false, progress.Visible, name + ": progress hidden");
        Equal(ProgressBarStyle.Blocks, progress.Style, name + ": progress style restored");
        Equal(0, progress.Value, name + ": progress value cleared");
        Equal(0, Field<DataGridView>(form, "sectionMetrics").Rows.Count, name + ": old EP rows cleared");
        Equal(false, Field<DataGridView>(form, "sectionMetrics").Visible, name + ": old EP results hidden");
        Equal(true, Field<Label>(form, "metrics").Visible && Field<Label>(form, "metrics").Text.Contains("Yellow insertions"), name + ": initial metrics guidance restored");
        Equal(true, Field<Label>(form, "state").Text.StartsWith("Ready"), name + ": initial ready status restored");
        Equal(true, Field<TextBox>(form, "details").Text.Contains("Results: crossover percentage"), name + ": initial result guidance restored");
        Equal(false, Field<Button>(form, "cancel").Enabled, name + ": cancellation unavailable when idle");
        Equal(true, Field<Panel>(form, "inputs").Enabled && Field<Button>(form, "compare").Enabled && Field<Button>(form, "reset").Enabled, name + ": new-job controls available");
    }

    private static string Hash(string path) { using (SHA256 hash = SHA256.Create()) using (FileStream source = File.OpenRead(path)) return Convert.ToBase64String(hash.ComputeHash(source)); }
    private static T Field<T>(object instance, string name) { return (T)instance.GetType().GetField(name, BindingFlags.Instance | BindingFlags.NonPublic).GetValue(instance); }
    private static void Set(object instance, string name, object value) { instance.GetType().GetField(name, BindingFlags.Instance | BindingFlags.NonPublic).SetValue(instance, value); }
    private static object Invoke(object instance, string name, params object[] arguments) { return instance.GetType().GetMethod(name, BindingFlags.Instance | BindingFlags.NonPublic).Invoke(instance, arguments); }
    private static T Named<T>(Control parent, string name) where T : Control { return (T)parent.Controls.Find(name, true)[0]; }
    private static void InBounds(Control control, Control parent, string name)
    {
        if (control.Left < 0 || control.Top < 0 || control.Right > parent.ClientSize.Width || control.Bottom > parent.ClientSize.Height)
            throw new Exception(name + ": control " + control.Bounds + "; parent " + parent.ClientSize);
        checks++;
    }
    private static void Equal<T>(T expected, T actual, string name)
    {
        if (!EqualityComparer<T>.Default.Equals(expected, actual)) throw new Exception(name + ": expected " + expected + ", got " + actual);
        checks++;
    }
}
