using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using LocalDocumentCompare;

internal static class LanguageMismatchUiTests
{
    private static int checks;
    private static string root;
    private delegate bool WindowCallback(IntPtr handle, IntPtr data);
    [DllImport("user32.dll")] private static extern bool EnumWindows(WindowCallback callback, IntPtr data);
    [DllImport("user32.dll")] private static extern bool EnumChildWindows(IntPtr parent, WindowCallback callback, IntPtr data);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern IntPtr SendMessageTimeout(IntPtr window, uint message, IntPtr wparam, StringBuilder text, uint flags, uint timeout, out IntPtr result);
    [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr window, out uint process);
    [DllImport("user32.dll")] private static extern IntPtr GetWindow(IntPtr window, uint command);
    [DllImport("user32.dll")] private static extern bool PostMessage(IntPtr window, uint message, IntPtr wparam, IntPtr lparam);

    [STAThread]
    public static int Main()
    {
        string parent = Path.GetFullPath(Path.Combine(Path.GetTempPath(), "LocalDocumentCompare")) + Path.DirectorySeparatorChar;
        root = Path.Combine(parent, "language-popup-ui-" + Guid.NewGuid().ToString("N")); Directory.CreateDirectory(root);
        var completed = new ManualResetEvent(false);
        Thread watchdog = new Thread(delegate() {
            if (!completed.WaitOne(60000)) { Console.Error.WriteLine("Owned synthetic popup harness exceeded 60 seconds."); Environment.Exit(2); }
        });
        watchdog.IsBackground = true; watchdog.Start();
        try
        {
            Application.EnableVisualStyles(); Application.SetCompatibleTextRenderingDefault(false);
            Single(); Batch();
            Console.WriteLine("PASS: " + checks + " language popup UI checks; actual owned Windows dialogs; cleanup and controls verified; no Word or real comparisons started.");
            return 0;
        }
        catch (Exception error) { Console.Error.WriteLine(error); return 1; }
        finally
        {
            completed.Set(); watchdog.Join(1000); completed.Dispose();
            string resolved = Path.GetFullPath(root);
            if (resolved.StartsWith(parent, StringComparison.OrdinalIgnoreCase) && Path.GetFileName(resolved).StartsWith("language-popup-ui-", StringComparison.Ordinal))
                try { Directory.Delete(resolved, true); } catch { }
        }
    }

    private static ComparisonResult Mismatch(string scope)
    {
        return new ComparisonResult { Success = false, Error = "Source language mismatch in " + scope + ": archive is German; new is English. Select documents in the same source language. No comparison has been created.",
            LanguageMismatch = new SourceLanguageMismatch { Scope = scope, ArchiveLanguage = "German", NewLanguage = "English" } };
    }
    private static void Single()
    {
        using (MainForm form = new MainForm { Opacity = 0, ShowInTaskbar = false })
        {
            form.Show(); Application.DoEvents();
            string work = SeedSingle(form, Mismatch("Description"), false);
            DialogWatch notice = Observe(form, LanguageMismatchNotice.SingleTitle, delegate {
                Equal(null, Field<Process>(form, "worker"), "single worker reference cleared before popup");
                Equal(false, Directory.Exists(work), "single temporary job removed before popup");
            });
            using (notice) Invoke(form, "FinishComparison");
            Equal(1, notice.Count, "one single failure popup");
            Equal(true, Field<Panel>(form, "inputs").Enabled && Field<Button>(form, "compare").Enabled, "single controls restored after popup");
            Equal(false, Field<Button>(form, "cancel").Enabled, "single Cancel disabled after popup");
            Equal(true, notice.Text.Contains("Description") && notice.Text.Contains("German") && notice.Text.Contains("English"), "visible popup explains section and languages");
            Equal(true, notice.Text.Contains("No comparison document has been created"), "visible popup explains no output");
            Equal(false, Field<Button>(form, "languageChecked").Visible, "known mismatch offers no manual override");
            Equal(true, Field<TextBox>(form, "details").Text.Contains("German"), "persistent details keep failure reason after popup");
            Field<Button>(form, "reset").PerformClick();
            Equal(false, Field<TextBox>(form, "details").Text.Contains("German"), "Reset clears old mismatch details");
            Equal(true, Field<Button>(form, "compare").Enabled, "single ready for next comparison after dismissal and Reset");

            foreach (ComparisonResult result in new[] { Mismatch("Description"), new ComparisonResult { Error = "The files could not be opened." },
                new ComparisonResult { Error = "A manual source-language check is required.", LanguageReviewRequired = true },
                new ComparisonResult { Success = true, NewWords = 100, InsertedWords = 20, CrossoverPercent = 80, OutputPath = "synthetic comparison.docx" } })
            {
                bool cancelled = result.LanguageMismatch != null;
                SeedSingle(form, result, cancelled);
                DialogWatch absent = Observe(form, LanguageMismatchNotice.SingleTitle, null);
                using (absent) { Invoke(form, "FinishComparison"); Application.DoEvents(); }
                Equal(0, absent.Count, cancelled ? "cancelled job has no stale mismatch popup" : "unrelated error/review/success has no mismatch popup");
                Field<Button>(form, "reset").PerformClick();
            }
            Equal(null, Field<Process>(form, "worker"), "single test never retains a worker");
            form.Close();
        }
    }
    private static string SeedSingle(MainForm form, ComparisonResult result, bool cancelled)
    {
        string work = Path.Combine(root, "single-job-" + Guid.NewGuid().ToString("N")); Directory.CreateDirectory(work);
        JsonFile.Write(Path.Combine(work, "result.json"), result);
        Set(form, "jobDirectory", work); Set(form, "worker", new Process()); Set(form, "cancelling", cancelled);
        Field<Panel>(form, "inputs").Enabled = false; Field<Button>(form, "compare").Enabled = false; Field<Button>(form, "cancel").Enabled = true;
        return work;
    }
    private static void Batch()
    {
        string inputs = Path.Combine(root, "batch-inputs"); Directory.CreateDirectory(inputs);
        var files = new List<BatchInput>();
        for (int i = 1; i <= 3; i++)
        {
            string path = Path.Combine(inputs, "0" + i + " synthetic.docx"); File.WriteAllText(path, "Synthetic placeholder, never opened in Word.");
            files.Add(new BatchInput { Path = path });
        }
        BatchPlan plan = BatchPlan.Create(files, root, "EP");
        var records = new List<BatchRecord> {
            new BatchRecord { Pair = plan.Pairs[0], State = "Failed", Result = Mismatch("Description") },
            new BatchRecord { Pair = plan.Pairs[1], State = "Failed", Result = Mismatch("Claims") },
            new BatchRecord { Pair = plan.Pairs[2], State = "Completed", Result = new ComparisonResult { Success = true, NewWords = 100, InsertedWords = 20, CrossoverPercent = 80 } }
        };
        string report = BatchReport.Save(plan, records);
        using (Form host = new Form { Opacity = 0, ShowInTaskbar = false, Width = 850, Height = 1000 })
        using (BatchComparisonControl batch = new BatchComparisonControl { Dock = DockStyle.Fill })
        {
            host.Controls.Add(batch); host.Show(); Application.DoEvents();
            SeedBatch(batch, plan, records, report);
            DialogWatch summary = Observe(host, LanguageMismatchNotice.BatchTitle, delegate {
                Equal(false, batch.IsBusy, "batch becomes idle before popup");
                Equal(3, Field<DataGridView>(batch, "resultsGrid").Rows.Count, "batch results retained before popup");
            });
            using (summary) Invoke(batch, "CompleteBatch", false, true);
            Equal(1, summary.Count, "one aggregate batch popup");
            Equal(true, Field<Panel>(batch, "configuration").Enabled, "batch controls restored after popup");
            Equal(true, Field<Button>(batch, "resetBatch").Enabled, "batch Reset available after popup");
            Equal(true, summary.Text.Contains("2 comparisons") && summary.Text.Contains("02 synthetic.docx") && summary.Text.Contains("03 synthetic.docx"), "batch popup identifies affected count and filenames");
            Equal(true, summary.Text.Contains("Description") && summary.Text.Contains("Claims") && summary.Text.Contains("Other completed comparisons are saved"), "batch popup explains scopes and retained successes");
            Equal(true, File.Exists(report), "batch report preserved after popup");
            DialogWatch duplicate = Observe(host, LanguageMismatchNotice.BatchTitle, null);
            using (duplicate) { Invoke(batch, "CompleteBatch", false, true); Invoke(batch, "SaveReport"); Application.DoEvents(); }
            Equal(0, duplicate.Count, "completion re-entry and report retry do not duplicate popup");
            Field<Button>(batch, "resetBatch").PerformClick();
            Equal(0, Field<DataGridView>(batch, "resultsGrid").Rows.Count, "batch Reset clears previous failure table");

            SeedBatch(batch, plan, records, report);
            DialogWatch closing = Observe(host, LanguageMismatchNotice.BatchTitle, null);
            using (closing) { Equal(true, batch.CancelAndClose(false), "batch can close with old mismatch results"); Application.DoEvents(); }
            Equal(0, closing.Count, "closing suppresses mismatch popup");
            Equal(false, batch.IsBusy, "closing leaves no active batch");
            Equal(null, Field<Process>(batch, "worker"), "no real batch worker was created");
            host.Close();
        }
    }
    private static void SeedBatch(BatchComparisonControl batch, BatchPlan plan, List<BatchRecord> records, string report)
    {
        Set(batch, "plan", plan); Set(batch, "records", records); Set(batch, "reportPath", report); Set(batch, "currentIndex", records.Count - 1);
        Set(batch, "suppressLanguageNotice", false); Set(batch, "cancelRequested", false); Invoke(batch, "SetBusy", true);
    }
    private static DialogWatch Observe(Form owner, string title, Action check) { return new DialogWatch(owner, title, check); }
    private sealed class DialogWatch : IDisposable
    {
        private readonly Thread watcher;
        private readonly IntPtr ownerHandle;
        private readonly string title;
        private readonly Action check;
        private readonly HashSet<IntPtr> seen = new HashSet<IntPtr>();
        private Exception failure;
        private volatile bool stopping;
        internal int Count;
        internal string Text = "";
        internal DialogWatch(Form owner, string title, Action check)
        {
            ownerHandle = owner.Handle; this.title = title; this.check = check;
            watcher = new Thread(delegate() { while (!stopping) { Scan(); Thread.Sleep(40); } });
            watcher.IsBackground = true; watcher.Start();
        }
        private void Scan()
        {
            uint ownPid = (uint)Process.GetCurrentProcess().Id;
            EnumWindows(delegate(IntPtr window, IntPtr unused) {
                uint pid; GetWindowThreadProcessId(window, out pid);
                if (pid != ownPid || WindowText(window) != title || GetWindow(window, 4) != ownerHandle) return true;
                if (seen.Add(window))
                {
                    Count++; Text += WindowText(window) + "\n";
                    EnumChildWindows(window, delegate(IntPtr child, IntPtr data) { Text += WindowText(child) + "\n"; return true; }, IntPtr.Zero);
                    try { if (check != null) check(); } catch (Exception error) { failure = error; }
                }
                // Dismiss only this process's dialog owned by the synthetic form.
                PostMessage(window, 0x0010, IntPtr.Zero, IntPtr.Zero); return true;
            }, IntPtr.Zero);
        }
        public void Dispose() { stopping = true; watcher.Join(2000); if (failure != null) throw failure; }
    }
    private static string WindowText(IntPtr window)
    {
        var text = new StringBuilder(12000); IntPtr copied;
        return SendMessageTimeout(window, 0x000D, new IntPtr(text.Capacity), text, 2, 200, out copied) == IntPtr.Zero ? "" : text.ToString();
    }
    private static T Field<T>(object instance, string name) { return (T)instance.GetType().GetField(name, BindingFlags.Instance | BindingFlags.NonPublic).GetValue(instance); }
    private static void Set(object instance, string name, object value) { instance.GetType().GetField(name, BindingFlags.Instance | BindingFlags.NonPublic).SetValue(instance, value); }
    private static object Invoke(object instance, string name, params object[] arguments)
    {
        foreach (MethodInfo method in instance.GetType().GetMethods(BindingFlags.Instance | BindingFlags.NonPublic))
            if (method.Name == name && method.GetParameters().Length == arguments.Length) return method.Invoke(instance, arguments);
        throw new Exception("Missing method " + name);
    }
    private static void Equal<T>(T expected, T actual, string label)
    {
        if (!EqualityComparer<T>.Default.Equals(expected, actual)) throw new Exception(label + ": expected " + expected + "; actual " + actual);
        checks++;
    }
}
