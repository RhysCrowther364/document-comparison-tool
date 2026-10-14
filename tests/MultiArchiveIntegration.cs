using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using LocalDocumentCompare;

internal static class MultiArchiveIntegration
{
    private const string English = "The present invention relates to a method for controlling a motor. The apparatus comprises a housing, a controller and a sensor. The controller receives the signal from the sensor and adjusts the speed of the motor. This arrangement improves the operation of the device. The following description explains how the components work together and how the operator can use the apparatus under different conditions. Each component performs a specific function within the system, and the controller records the result before the next operation begins.";
    private const string German = "Die vorliegende Erfindung betrifft ein Verfahren zur Steuerung eines Motors. Die Vorrichtung besteht aus einer Steuerung und einem Sensor. Die Steuerung verarbeitet das Signal des Sensors und stellt die Drehzahl des Motors ein. Durch diese Anordnung wird der Betrieb der Vorrichtung verbessert. Das Verfahren wird mit einer weiteren Einheit verbunden, damit der Motor unter verschiedenen Bedingungen sicher arbeiten kann. Die Beschreibung zeigt, wie die einzelnen Bauteile zusammenwirken und welche Funktionen bei der Anwendung des Verfahrens erforderlich sind.";
    private static Assembly app;
    private static string root;
    private static int checks;
    private static readonly string LongEnglish = String.Join(" ", Enumerable.Repeat(English, 8));
    private static readonly string LongGerman = String.Join(" ", Enumerable.Repeat(German, 8));

    [STAThread]
    public static int Main(string[] args)
    {
        root = Path.Combine(Path.GetTempPath(), "LocalCompare-MultiNative-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        try
        {
            app = Assembly.LoadFrom(Path.GetFullPath(args[0]));
            Application.EnableVisualStyles(); Application.SetCompatibleTextRenderingDefault(false);
            WfMatches(); EpMatches(); ManualReview(); CancelPartial();
            Console.WriteLine("PASS: " + checks + " native Multi Archive/Word checks; synthetic sources preserved and private jobs closed.");
            return 0;
        }
        catch (Exception ex) { Console.Error.WriteLine(ex); return 1; }
        finally
        {
            string resolved = Path.GetFullPath(root), boundary = Path.GetFullPath(Path.GetTempPath()).TrimEnd(Path.DirectorySeparatorChar) + Path.DirectorySeparatorChar;
            if (!resolved.StartsWith(boundary, StringComparison.OrdinalIgnoreCase) || !Path.GetFileName(resolved).StartsWith("LocalCompare-MultiNative-", StringComparison.Ordinal)) throw new Exception("Unexpected native test cleanup path.");
            Directory.Delete(resolved, true);
        }
    }

    private static void WfMatches()
    {
        string folder = Path.Combine(root, "wf"); Directory.CreateDirectory(folder);
        string revised = Doc(folder, "New", LongEnglish, "The final review contains sapphire amber coral.");
        string bad = Doc(folder, "German", LongGerman);
        string near = Doc(folder, "Near", LongEnglish, "The final review contains sapphire amber.");
        string exact = Doc(folder, "Exact", LongEnglish, "The final review contains sapphire amber coral.");
        string lesser = Doc(folder, "Lesser", LongEnglish, "The final review contains sapphire.");
        string[] paths = { bad, near, exact, lesser, revised };
        var hashes = paths.ToDictionary(p => p, Hash);
        using (Session session = new Session())
        using (DialogWatch watch = new DialogWatch(session.Form))
        {
            session.Setup(revised, paths.Take(4).ToArray(), false, folder);
            session.Click("startBatch");
            Pump(() => !session.Busy, "WF matching and mismatch continuation");
            List<BatchRecord> records = session.Records;
            Require(records.Count == 4, "Four archives must produce exactly four comparisons.");
            Require(records[0].State == "Failed" && records[0].Result.LanguageMismatch != null, "Known mismatch must fail individually.");
            Require(records.Skip(1).All(r => r.State == "Completed" && r.Result.Success), "Remaining archive matches did not continue.");
            int words = DocxContent.Analyze(revised).Words;
            CheckResult(records[1], words, 1); CheckResult(records[2], words, 0); CheckResult(records[3], words, 2);
            Require(MultiArchiveRanking.Ordered(records, "WF")[0].Pair.Archive.Path == exact, "Exact archive not ranked first.");
            DataGridView grid = (DataGridView)Field(session.Control, "resultsGrid");
            Require((string)grid.Rows[0].Cells["archive"].Value == "Exact.docx" && (string)grid.Rows[0].Cells["rank"].Value == "1", "Actual UI ranking incorrect.");
            Require(((TextBox)Field(session.Control, "details")).Text.Contains("Best match among ranked archives: Exact.docx — 100.0%"), "Missing best-match summary.");
            string[] csv = File.ReadAllLines(session.Plan.ReportPath);
            Require(csv.Length == 5 && csv[1].StartsWith("\"1\",\"3\""), "Ranked CSV differs from actual results.");
            Require(csv[4].StartsWith("\"\",\"1\"") && csv[4].Contains("Failed"), "Mismatch must be unranked in CSV.");
            Require(watch.Count == 1 && watch.Text.Contains("English") && watch.Text.Contains("German"), "One clear Multi Archive mismatch summary expected.");
            AssertFinished(session);
            foreach (string path in paths) Require(hashes[path] == Hash(path), "WF source changed.");
            using (Bitmap image = new Bitmap(session.Form.Width, session.Form.Height)) { session.Form.DrawToBitmap(image, new Rectangle(Point.Empty, session.Form.Size)); image.Save(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "multi-archive-results-preview.png")); }
        }
        Console.WriteLine("PASS native Full Document: one new vs four archives, ranked matches, mismatch continuation, CSV, highlights and source hashes.");
    }

    private static void EpMatches()
    {
        string folder = Path.Combine(root, "ep"); Directory.CreateDirectory(folder);
        string newDescription = LongEnglish + " Sapphire amber coral.";
        string newClaim = "1. " + LongEnglish + " The device contains sapphire amber coral.";
        string revised = Doc(folder, "New", "Description", newDescription, "Claims", newClaim);
        string a = Doc(folder, "DescriptionBest", "Description", newDescription, "Claims", "1. " + LongEnglish + " The device contains sapphire.", "Anspr\u00fcche", "1. TRANSLATIONEXCLUDE");
        string b = Doc(folder, "SpecificationBest", "Description", LongEnglish + " Sapphire amber.", "Claims", newClaim);
        string[] paths = { revised, a, b }; var hashes = paths.ToDictionary(p => p, Hash);
        using (Session session = new Session())
        {
            session.Setup(revised, new[] { a, b }, true, folder);
            session.Click("startBatch"); Pump(() => !session.Busy, "EP archive matching");
            List<BatchRecord> records = session.Records;
            Require(records.All(r => r.State == "Completed" && r.Result.Sections.Length == 3), "EP matches must return all three scopes.");
            Require(records[0].Result.Sections[0].InsertedWords == 0 && records[0].Result.Sections[2].InsertedWords == 2, "First EP match scope counts differ.");
            Require(records[1].Result.Sections[0].InsertedWords == 1 && records[1].Result.Sections[2].InsertedWords == 0, "Second EP match scope counts differ.");
            Require(MultiArchiveRanking.Ordered(records, "EP")[0].Pair.Archive.Path == b, "EP must rank Full specification, not Description.");
            foreach (BatchRecord record in records)
            {
                Require(record.Result.Sections[0].NewWords + record.Result.Sections[2].NewWords == record.Result.Sections[1].NewWords, "EP word sums differ.");
                Require(record.Result.Sections.All(s => s.CrossoverPercent.HasValue), "EP strict reconciliation withheld a test percentage.");
                Require(DocxContent.ReadFinalText(record.Pair.OutputPath).IndexOf("TRANSLATIONEXCLUDE", StringComparison.Ordinal) < 0, "Translated claims leaked into output.");
                CheckYellow(record.Pair.OutputPath);
            }
            DataGridView grid = (DataGridView)Field(session.Control, "resultsGrid");
            Require(grid.Rows.Count == 6 && (string)grid.Rows[0].Cells["archive"].Value == "SpecificationBest.docx", "EP ranked grid incorrect.");
            string[] csv = File.ReadAllLines(session.Plan.ReportPath);
            Require(csv.Length == 7 && csv.Skip(1).Take(3).All(l => l.StartsWith("\"1\",\"2\"")), "EP CSV must repeat rank for each of three scopes.");
            AssertFinished(session);
            foreach (string path in paths) Require(hashes[path] == Hash(path), "EP source changed.");
        }
        Console.WriteLine("PASS native EP: Full specification ranking, three scopes, claims exclusion, strict figures and yellow insertions.");
    }

    private static void ManualReview()
    {
        string folder = Path.Combine(root, "review"); Directory.CreateDirectory(folder);
        string archive = Doc(folder, "Archive", "Alpha beta gamma delta.");
        string revised = Doc(folder, "New", "Alpha beta gamma delta epsilon.");
        using (Session session = new Session())
        {
            session.Setup(revised, new[] { archive }, false, folder);
            session.Click("startBatch");
            Pump(() => session.Busy && (bool)Field(session.Control, "reviewPaused"), "manual language review");
            Require(session.Records[0].State == "Review required", "Uncertain language must stop this archive.");
            Require(!File.Exists(session.Plan.Pairs[0].OutputPath), "Unconfirmed language published comparison.");
            Require(MultiArchiveRanking.Ranks(session.Records, "WF").Count == 0, "Unconfirmed pair was ranked.");
            session.Click("languageChecked"); Pump(() => !session.Busy, "confirmed current-archive retry");
            CheckResult(session.Records[0], 5, 1);
            Require(session.Records[0].Result.CrossoverPercent == 80, "Manual retry figures differ.");
            Require(MultiArchiveRanking.Ranks(session.Records, "WF")[1] == 1, "Confirmed result not ranked.");
            AssertFinished(session);
        }
        Console.WriteLine("PASS native manual review: no premature output or rank; explicit current-pair retry returns 80.0% / 5 / 1.");
    }

    private static void CancelPartial()
    {
        string folder = Path.Combine(root, "cancel"); Directory.CreateDirectory(folder);
        string revised = Doc(folder, "New", LongEnglish);
        string[] archives = Enumerable.Range(1, 3).Select(i => Doc(folder, "Archive" + i, LongEnglish)).ToArray();
        using (Session session = new Session())
        {
            session.Setup(revised, archives, false, folder);
            session.Click("startBatch");
            Pump(() => session.Records.Count == 3 && session.Records[0].State == "Completed" && Field(session.Control, "worker") != null, "first result before cancellation");
            string completed = session.Records[0].Pair.OutputPath, before = Hash(completed);
            session.Click("cancelBatch"); Pump(() => !session.Busy, "archive search cancellation");
            Require(File.Exists(completed) && Hash(completed) == before, "Cancellation changed a completed comparison.");
            Require(session.Records.Any(r => r.State == "Not run") && session.Records.All(r => r.State != "Pending" && r.State != "Running"), "Cancellation statuses incomplete.");
            Require(MultiArchiveRanking.Ranks(session.Records, "WF").Count == session.Records.Count(r => r.State == "Completed"), "Cancelled results acquired a rank.");
            Require(((TextBox)Field(session.Control, "details")).Text.Contains("among completed comparisons"), "Partial best-match summary is misleading.");
            AssertFinished(session);
        }
        Console.WriteLine("PASS native cancellation: completed comparison and partial ranking retained; unfinished archives unranked.");
    }

    private static void CheckResult(BatchRecord record, int words, int inserted)
    {
        Require(record.State == "Completed" && record.Result.Success, "Expected completed comparison.");
        Require(record.Result.NewWords == words && record.Result.InsertedWords == inserted, "Native counts differ.");
        Require(record.Result.CrossoverPercent.HasValue && Math.Abs(record.Result.CrossoverPercent.Value - 100.0 * (words - inserted) / words) < 0.000001, "Native crossover differs.");
        Require(File.Exists(record.Pair.OutputPath) && File.Exists(Path.ChangeExtension(record.Pair.OutputPath, ".txt")), "Comparison or summary missing.");
        if (inserted > 0) CheckYellow(record.Pair.OutputPath);
    }
    private static void CheckYellow(string path)
    {
        using (ZipArchive zip = ZipFile.OpenRead(path))
        using (StreamReader reader = new StreamReader(zip.GetEntry("word/document.xml").Open()))
        { string xml = reader.ReadToEnd(); Require(xml.Contains("w:ins") && xml.Contains("w:val=\"yellow\""), "Actual inserted words lack yellow tracked markup."); }
    }
    private static void AssertFinished(Session session)
    {
        Require(Field(session.Control, "worker") == null && Field(session.Control, "jobDirectory") == null && !(bool)Field(session.Control, "keepJobDirectory"), "Private job or worker left behind.");
        Require(((TabPage)Field(session.Form, "singleTab")).Enabled && ((TabPage)Field(session.Form, "batchTab")).Enabled, "Other tabs did not unlock.");
    }
    private sealed class Session : IDisposable
    {
        internal readonly Form Form;
        internal readonly object Control;
        internal Session()
        {
            Form = (Form)Activator.CreateInstance(app.GetType("LocalDocumentCompare.MainForm"), true);
            Form.Opacity = 0; Form.ShowInTaskbar = false; Form.Show();
            Control = Field(Form, "multiArchive");
            ((TabControl)Field(Form, "tabs")).SelectedTab = (TabPage)Field(Form, "multiArchiveTab"); Application.DoEvents();
        }
        internal bool Busy { get { return (bool)Control.GetType().GetProperty("IsBusy").GetValue(Control, null); } }
        internal BatchPlan Plan { get { return (BatchPlan)Field(Control, "plan"); } }
        internal List<BatchRecord> Records { get { return (List<BatchRecord>)Field(Control, "records"); } }
        internal void Setup(string revised, string[] archives, bool ep, string output)
        {
            ((TextBox)Field(Control, "newDocument")).Text = revised;
            Invoke(Control, "AddArchivePaths", (object)archives);
            ((ComboBox)Field(Control, "jobType")).SelectedIndex = ep ? 1 : 0;
            ((TextBox)Field(Control, "outputPath")).Text = output;
        }
        internal void Click(string name) { ((Button)Field(Control, name)).PerformClick(); Application.DoEvents(); }
        public void Dispose()
        {
            if (Busy) Control.GetType().GetMethod("CancelAndClose", BindingFlags.Instance | BindingFlags.NonPublic).Invoke(Control, new object[] { false });
            Form.Dispose();
        }
    }
    private static object Field(object target, string name) { return target.GetType().GetField(name, BindingFlags.Instance | BindingFlags.NonPublic).GetValue(target); }
    private static object Invoke(object target, string name, params object[] args) { return target.GetType().GetMethod(name, BindingFlags.Instance | BindingFlags.NonPublic).Invoke(target, args); }
    private static void Pump(Func<bool> condition, string label)
    {
        DateTime end = DateTime.UtcNow.AddMinutes(4);
        while (!condition()) { Application.DoEvents(); if (DateTime.UtcNow >= end) throw new Exception("Timed out: " + label); Thread.Sleep(30); }
        Application.DoEvents();
    }
    private static string Doc(string folder, string stem, params string[] paragraphs)
    { string path = Path.Combine(folder, stem + ".docx"); DocxContent.CreateFromRecognizedText(path, String.Join("\n", paragraphs)); return path; }
    private static string Hash(string path) { using (var hash = SHA256.Create()) using (var stream = File.OpenRead(path)) return BitConverter.ToString(hash.ComputeHash(stream)); }
    private static void Require(bool condition, string message) { if (!condition) throw new Exception(message); checks++; }

    private delegate bool WindowCallback(IntPtr window, IntPtr data);
    [DllImport("user32.dll")] private static extern bool EnumWindows(WindowCallback callback, IntPtr data);
    [DllImport("user32.dll")] private static extern bool EnumChildWindows(IntPtr window, WindowCallback callback, IntPtr data);
    [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr window, out uint pid);
    [DllImport("user32.dll")] private static extern IntPtr GetWindow(IntPtr window, uint command);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern IntPtr SendMessageTimeout(IntPtr window, uint message, IntPtr wparam, StringBuilder text, uint flags, uint timeout, out IntPtr result);
    [DllImport("user32.dll")] private static extern bool PostMessage(IntPtr window, uint message, IntPtr wparam, IntPtr lparam);
    private sealed class DialogWatch : IDisposable
    {
        private readonly Thread thread;
        private readonly IntPtr owner;
        private volatile bool stopping;
        internal int Count;
        internal string Text = "";
        internal DialogWatch(Form form)
        {
            owner = form.Handle;
            thread = new Thread(delegate() {
                var seen = new HashSet<IntPtr>(); uint ownPid = (uint)Process.GetCurrentProcess().Id;
                while (!stopping)
                {
                    EnumWindows(delegate(IntPtr window, IntPtr unused) {
                        uint pid; GetWindowThreadProcessId(window, out pid);
                        if (pid != ownPid || GetWindow(window, 4) != owner || WindowText(window) != "Multi Archive — source language mismatches") return true;
                        if (seen.Add(window)) { Count++; EnumChildWindows(window, delegate(IntPtr child, IntPtr data) { Text += WindowText(child) + "\n"; return true; }, IntPtr.Zero); }
                        PostMessage(window, 0x0010, IntPtr.Zero, IntPtr.Zero); return true;
                    }, IntPtr.Zero);
                    Thread.Sleep(40);
                }
            });
            thread.IsBackground = true; thread.Start();
        }
        public void Dispose() { stopping = true; thread.Join(2000); }
    }
    private static string WindowText(IntPtr window)
    { var text = new StringBuilder(6000); IntPtr copied; return SendMessageTimeout(window, 0x000D, new IntPtr(text.Capacity), text, 2, 200, out copied) == IntPtr.Zero ? "" : text.ToString(); }
}
