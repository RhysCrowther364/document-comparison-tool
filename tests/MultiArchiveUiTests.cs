using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Windows.Forms;
using LocalDocumentCompare;

internal static class MultiArchiveUiTests
{
    private static int checks;
    [STAThread]
    public static int Main()
    {
        string root = Path.Combine(Path.GetTempPath(), "LocalCompare-MultiUi-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        try
        {
            Application.EnableVisualStyles(); Application.SetCompatibleTextRenderingDefault(false);
            string newPath = Path.Combine(root, "New.docx"), pdfPath = Path.Combine(root, "Scanned.pdf");
            string a = Path.Combine(root, "Archive1.docx"), b = Path.Combine(root, "Archive2.docx");
            foreach (string path in new[] { newPath, pdfPath, a, b }) File.WriteAllText(path, "synthetic placeholder; never opened");
            using (MainForm form = new MainForm { Opacity = 0, ShowInTaskbar = false })
            {
                form.Show(); Application.DoEvents();
                TabControl tabs = Field<TabControl>(form, "tabs");
                BatchComparisonControl multi = Field<BatchComparisonControl>(form, "multiArchive");
                BatchComparisonControl batch = Field<BatchComparisonControl>(form, "batch");
                tabs.SelectedTab = Field<TabPage>(form, "multiArchiveTab"); Application.DoEvents();
                Equal(3, tabs.TabPages.Count, "three comparison tabs");
                Equal("Multi Archive", tabs.SelectedTab.Text, "requested tab label");
                TextBox revised = Field<TextBox>(multi, "newDocument");
                CheckBox ocr = Field<CheckBox>(multi, "newDocumentOcr");
                Button start = Field<Button>(multi, "startBatch");
                List<BatchInput> files = Field<List<BatchInput>>(multi, "files");
                Equal(false, start.Enabled, "empty search cannot start");
                Equal(false, ocr.Enabled, "new OCR starts unavailable");
                Equal(true, Field<Button>(multi, "addArchives").Visible, "archive file picker visible");
                Equal(false, Field<Button>(multi, "moveUp").Visible, "all-pairs ordering controls hidden");
                Equal(true, Field<DataGridView>(multi, "fileGrid").AllowDrop, "archive list accepts drops");
                Equal(true, revised.AllowDrop, "new field accepts drops");
                revised.Text = pdfPath; ocr.Checked = true;
                Equal(true, ocr.Enabled, "new PDF OCR can be selected");
                revised.Text = newPath;
                Equal(false, ocr.Enabled, "Word new input cannot select OCR");
                Equal(false, ocr.Checked, "new-file change clears OCR choice");
                Invoke(multi, "AddArchivePaths", (object)new[] { a, pdfPath });
                Equal(2, files.Count, "multiple archive files can be added");
                Equal(true, start.Enabled, "one new and archives enable search");
                DataGridView inputGrid = Field<DataGridView>(multi, "fileGrid");
                inputGrid.Rows[1].Cells[2].Value = true;
                Invoke(multi, "AddArchivePaths", (object)new[] { pdfPath, b });
                Equal(3, files.Count, "duplicate additions are ignored");
                Equal(true, files[1].UseOcr, "duplicate add preserves selected PDF OCR");
                Equal("3 archives · 3 comparisons", Field<Label>(multi, "pairCount").Text, "one comparison per archive rather than all pairs");
                string invalid = Path.Combine(root, "invalid.txt"); File.WriteAllText(invalid, "synthetic");
                Invoke(multi, "AddArchivePaths", (object)new[] { a, invalid });
                Equal(3, files.Count, "invalid add does not partly change selection");
                Field<TextBox>(multi, "folderPath").Text = root;
                Invoke(multi, "LoadFiles", null, EventArgs.Empty);
                Equal(3, files.Count, "folder loading excludes selected new file");

                ComboBox mode = Field<ComboBox>(multi, "jobType"); mode.SelectedIndex = 1;
                revised.Text = pdfPath;
                Equal(true, Field<Label>(multi, "jobHint").Text.Contains("new file is a PDF"), "EP PDF recommendation uses the fixed new input");
                revised.Text = newPath; mode.SelectedIndex = 0;
                foreach (Size size in new[] { new Size(820, 864), new Size(760, 580) })
                {
                    form.Size = size; Application.DoEvents();
                    Panel config = Field<Panel>(multi, "configuration");
                    foreach (string name in new[] { "browseMultiNew", "browseFolder", "browseOutput", "saveDesktop" })
                    {
                        Control button = config.Controls.Find(name, false)[0];
                        Equal(true, button.Left >= 0 && button.Right <= config.ClientSize.Width, name + " stays within width " + size.Width);
                    }
                    Equal(true, revised.Right <= config.ClientSize.Width, "new input stays within configuration width");
                }
                string output = Path.Combine(root, "results"); Field<TextBox>(multi, "outputPath").Text = output;
                BatchPlan plan = BatchPlan.CreateMultiArchive(new[] { new BatchInput { Path = a }, new BatchInput { Path = b } }, new BatchInput { Path = newPath }, output, "WF");
                var records = new List<BatchRecord> {
                    Done(plan.Pairs[0], 4, 60), Done(plan.Pairs[1], 1, 90)
                };
                Set(multi, "plan", plan); Set(multi, "records", records);
                Invoke(multi, "SetBusy", true);
                Equal(false, Field<TabPage>(form, "singleTab").Enabled, "search locks single tab");
                Equal(false, Field<TabPage>(form, "batchTab").Enabled, "search locks all-pairs tab");
                Equal(false, Field<Button>(multi, "resetBatch").Enabled, "Reset disabled during search");
                tabs.SelectedTab = Field<TabPage>(form, "singleTab");
                Equal("Multi Archive", tabs.SelectedTab.Text, "cannot switch away from active search");
                Field<TextBox>(form, "archive").Text = a;
                Invoke(form, "StartComparison", null, EventArgs.Empty);
                Equal(null, Field<object>(form, "worker"), "single comparison cannot start while search is active");
                Invoke(multi, "ResetBatch"); Equal(newPath, revised.Text, "active reset leaves new selection unchanged");
                Invoke(multi, "RefreshResults");
                DataGridView grid = Field<DataGridView>(multi, "resultsGrid");
                Equal("Rank so far", grid.Columns["rank"].HeaderText, "running ranks marked provisional");
                Equal("Archive2.docx", grid.Rows[0].Cells["archive"].Value, "best archive shown first");
                Equal("1", grid.Rows[0].Cells["rank"].Value, "best rank visible");
                Equal("New.docx", grid.Rows[0].Cells["new"].Value, "fixed new document shown");
                Invoke(multi, "CompleteBatch", false, false);
                Equal("Rank", grid.Columns["rank"].HeaderText, "final rank header");
                Equal(true, Field<TextBox>(multi, "details").Text.Contains("Best match among ranked archives: Archive2.docx — 90.0%"), "clear best-match summary");
                Equal(true, Field<TabPage>(form, "singleTab").Enabled, "single tab restored after search");
                Equal(true, Field<TabPage>(form, "batchTab").Enabled, "batch tab restored after search");
                Invoke(batch, "SetBusy", true);
                Equal(false, Field<TabPage>(form, "multiArchiveTab").Enabled, "batch activity locks search tab");
                Invoke(batch, "SetBusy", false);
                Equal(true, Field<TabPage>(form, "multiArchiveTab").Enabled, "search restored after batch");

                string report = BatchReport.Save(plan, records); Set(multi, "reportPath", report);
                string reportBefore = File.ReadAllText(report);
                mode.SelectedIndex = 1;
                Invoke(multi, "ResetBatch");
                Equal("", revised.Text, "Reset clears fixed new file");
                Equal(false, ocr.Checked, "Reset clears new OCR");
                Equal(0, files.Count, "Reset clears all archives");
                Equal(0, grid.Rows.Count, "Reset clears ranked results");
                Equal(null, Field<BatchPlan>(multi, "plan"), "Reset discards prior plan reference");
                Equal(1, mode.SelectedIndex, "Reset preserves EP job type");
                Equal(output, Field<TextBox>(multi, "outputPath").Text, "Reset preserves destination");
                Equal(reportBefore, File.ReadAllText(report), "Reset preserves saved ranked report");
                Equal(false, start.Enabled, "Reset requires fresh files");
                using (Bitmap image = new Bitmap(form.Width, form.Height)) { form.DrawToBitmap(image, new Rectangle(Point.Empty, form.Size)); image.Save(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "multi-archive-preview.png")); }
            }
            Console.WriteLine("PASS: " + checks + " Multi Archive UI, tab locking, selection, ranking, layout and Reset checks.");
            return 0;
        }
        catch (Exception ex) { Console.Error.WriteLine(ex); return 1; }
        finally
        {
            string resolved = Path.GetFullPath(root), boundary = Path.GetFullPath(Path.GetTempPath()).TrimEnd(Path.DirectorySeparatorChar) + Path.DirectorySeparatorChar;
            if (!resolved.StartsWith(boundary, StringComparison.OrdinalIgnoreCase) || !Path.GetFileName(resolved).StartsWith("LocalCompare-MultiUi-", StringComparison.Ordinal)) throw new Exception("Unexpected UI test cleanup path.");
            Directory.Delete(resolved, true);
        }
    }
    private static BatchRecord Done(BatchPair pair, int inserted, double crossover) { return new BatchRecord { Pair = pair, State = "Completed", Result = new ComparisonResult { Success = true, NewWords = 10, InsertedWords = inserted, CrossoverPercent = crossover } }; }
    private static T Field<T>(object target, string name) { return (T)target.GetType().GetField(name, BindingFlags.Instance | BindingFlags.NonPublic).GetValue(target); }
    private static void Set(object target, string name, object value) { target.GetType().GetField(name, BindingFlags.Instance | BindingFlags.NonPublic).SetValue(target, value); }
    private static object Invoke(object target, string name, params object[] args) { return target.GetType().GetMethod(name, BindingFlags.Instance | BindingFlags.NonPublic).Invoke(target, args); }
    private static void Equal(object expected, object actual, string label) { if (!Object.Equals(expected, actual)) throw new Exception(label + ": expected " + expected + "; got " + actual); checks++; }
}
