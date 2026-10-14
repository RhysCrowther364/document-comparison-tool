using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Reflection;
using System.Security;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using LocalDocumentCompare;

// Actual WinForms controls and the shipped worker, using synthetic documents only.
// Run separately from other Word automation tests. Artifacts stay local for review.
internal static class BatchUiIntegration
{
    private const string English = "The present invention relates to a method for controlling a motor. The apparatus comprises a housing, a controller and a sensor. The controller receives the signal from the sensor and adjusts the speed of the motor. This arrangement improves the operation of the device. The following description explains how the components work together and how the operator can use the apparatus under different conditions. Each component performs a specific function within the system, and the controller records the result before the next operation begins.";
    private const string German = "Die vorliegende Erfindung betrifft ein Verfahren zur Steuerung eines Motors. Die Vorrichtung besteht aus einer Steuerung und einem Sensor. Die Steuerung verarbeitet das Signal des Sensors und stellt die Drehzahl des Motors ein. Durch diese Anordnung wird der Betrieb der Vorrichtung verbessert. Das Verfahren wird mit einer weiteren Einheit verbunden, damit der Motor unter verschiedenen Bedingungen sicher arbeiten kann. Die Beschreibung zeigt, wie die einzelnen Bauteile zusammenwirken und welche Funktionen bei der Anwendung des Verfahrens erforderlich sind.";
    private static Assembly app;
    private static string root;
    private static int checks;
    private static readonly string LongEnglish = String.Join(" ", Enumerable.Repeat(English, 8));
    private static readonly string LongGerman = String.Join(" ", Enumerable.Repeat(German, 8));

    private static object Field(object target, string name)
    {
        return target.GetType().GetField(name, BindingFlags.Instance | BindingFlags.NonPublic).GetValue(target);
    }
    private static object Property(object target, string name)
    {
        return target.GetType().GetProperty(name).GetValue(target, null);
    }
    private static void Require(bool condition, string message)
    {
        if (!condition) throw new InvalidOperationException(message);
        checks++;
    }
    private static string Hash(string path)
    {
        using (SHA256 algorithm = SHA256.Create())
        using (FileStream stream = File.OpenRead(path))
            return BitConverter.ToString(algorithm.ComputeHash(stream));
    }
    private static string NewDoc(string folder, string name, params string[] paragraphs)
    {
        Directory.CreateDirectory(folder);
        string path = Path.Combine(folder, name + ".docx");
        using (ZipArchive zip = ZipFile.Open(path, ZipArchiveMode.Create))
        {
            Add(zip, "[Content_Types].xml", "<Types xmlns='http://schemas.openxmlformats.org/package/2006/content-types'><Default Extension='rels' ContentType='application/vnd.openxmlformats-package.relationships+xml'/><Default Extension='xml' ContentType='application/xml'/><Override PartName='/word/document.xml' ContentType='application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml'/></Types>");
            Add(zip, "_rels/.rels", "<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'><Relationship Id='rId1' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument' Target='word/document.xml'/></Relationships>");
            string body = String.Join("", paragraphs.Select(p => "<w:p><w:r><w:t xml:space='preserve'>" + SecurityElement.Escape(p) + "</w:t></w:r></w:p>"));
            Add(zip, "word/document.xml", "<w:document xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'><w:body>" + body + "<w:sectPr/></w:body></w:document>");
        }
        return path;
    }
    private static void Add(ZipArchive zip, string name, string text)
    {
        using (StreamWriter writer = new StreamWriter(zip.CreateEntry(name).Open(), new UTF8Encoding(false))) writer.Write(text);
    }
    private static void PumpUntil(Func<bool> condition, string reason)
    {
        DateTime end = DateTime.UtcNow.AddMinutes(12);
        while (!condition())
        {
            Application.DoEvents();
            if (DateTime.UtcNow > end) throw new InvalidOperationException("Timed out: " + reason);
            Thread.Sleep(20);
        }
        Application.DoEvents();
    }
    private static bool Busy(object batch) { return (bool)Property(batch, "IsBusy"); }
    private static IList Records(object batch) { return (IList)Field(batch, "records"); }
    private static string State(object record) { return (string)Property(record, "State"); }
    private static ComparisonResult Result(object record) { return (ComparisonResult)Property(record, "Result"); }
    private static void Click(object batch, string name) { ((Button)Field(batch, name)).PerformClick(); Application.DoEvents(); }
    private static void Save(Form form, string name)
    {
        Application.DoEvents();
        using (Bitmap image = new Bitmap(form.Width, form.Height))
        {
            form.DrawToBitmap(image, new Rectangle(Point.Empty, form.Size));
            image.Save(Path.Combine(root, name), System.Drawing.Imaging.ImageFormat.Png);
        }
    }
    private static void WithForm(string input, string name, bool ep, Action<Form, object> test)
    {
        Dictionary<string, string> hashes = Directory.GetFiles(input, "*.docx").ToDictionary(p => p, p => Hash(p));
        using (Form form = (Form)Activator.CreateInstance(app.GetType("LocalDocumentCompare.MainForm", true)))
        {
            object batch = Field(form, "batch");
            try
            {
                form.Opacity = 0;
                form.ShowInTaskbar = false;
                form.StartPosition = FormStartPosition.Manual;
                form.Location = new Point(60, 60);
                form.Show();
                ((TabControl)Field(form, "tabs")).SelectedIndex = 1;
                ((TextBox)Field(batch, "folderPath")).Text = input;
                ((TextBox)Field(batch, "outputPath")).Text = Path.Combine(root, name + "-results");
                ((ComboBox)Field(batch, "jobType")).SelectedIndex = ep ? 1 : 0;
                Click(batch, "loadFiles");
                test(form, batch);
                foreach (KeyValuePair<string, string> source in hashes)
                    Require(Hash(source.Key) == source.Value, name + ": source changed: " + source.Key);
            }
            finally
            {
                if (Busy(batch)) batch.GetType().GetMethod("CancelAndClose").Invoke(batch, null);
                form.Hide();
            }
        }
    }
    private static string Output(object batch) { return (string)Property(Field(batch, "plan"), "OutputDirectory"); }
    private static void Finished(object batch)
    {
        Require(!Busy(batch) && Field(batch, "worker") == null, "Batch stayed busy after completion.");
        Require(Field(batch, "jobDirectory") == null, "Batch retained a completed temporary job.");
        Require(((Button)Field(batch, "startBatch")).Enabled && !((Button)Field(batch, "cancelBatch")).Enabled, "Batch actions did not reset.");
        Require(File.Exists(Path.Combine(Output(batch), "batch-results.csv")), "Missing combined CSV.");
    }
    private static void CheckOutput(object record)
    {
        ComparisonResult result = Result(record);
        Require(result != null && result.Success && File.Exists(result.OutputPath) && File.Exists(result.SummaryPath), "Completed pair lacks document or summary.");
        Require(DocxContent.Analyze(result.OutputPath).Words == result.NewWords, "Saved comparison word count differs from metrics.");
        using (ZipArchive zip = ZipFile.OpenRead(result.OutputPath))
        using (StreamReader reader = new StreamReader(zip.GetEntry("word/document.xml").Open()))
        {
            string xml = reader.ReadToEnd();
            if (result.InsertedWords > 0) Require(xml.Contains("yellow"), "Insertions were not highlighted yellow.");
        }
    }
    private static void SixPairs()
    {
        string input = Path.Combine(root, "four-inputs");
        NewDoc(input, "01", LongEnglish);
        NewDoc(input, "02", LongEnglish, "Amber.");
        NewDoc(input, "03", LongEnglish, "Amber cobalt.");
        NewDoc(input, "04", LongEnglish, "Amber cobalt jade.");
        File.WriteAllText(Path.Combine(input, "ignored.txt"), "Ignored");
        File.WriteAllText(Path.Combine(input, "~$ignored.docx"), "Word lock");
        NewDoc(Path.Combine(input, "subfolder"), "99", "Should not be loaded");
        int baseWords = DocxContent.Analyze(Path.Combine(input, "01.docx")).Words;
        WithForm(input, "six", false, delegate(Form form, object batch)
        {
            DataGridView files = (DataGridView)Field(batch, "fileGrid");
            Require(files.Rows.Count == 4, "Folder included unsupported, locked or nested files.");
            files.ClearSelection(); files.Rows[1].Selected = true; files.CurrentCell = files.Rows[1].Cells[0];
            Click(batch, "moveUp"); Click(batch, "moveDown");
            Save(form, "batch-ready.png");
            form.Size = new Size(760, 580);
            Save(form, "batch-minimum.png");
            form.ClientSize = new Size(820, Math.Min(864, Screen.PrimaryScreen.WorkingArea.Height - 90));
            Click(batch, "startBatch");
            Require(Busy(batch), "Batch did not start.");
            ((TabControl)Field(form, "tabs")).SelectedIndex = 0;
            Require(((TabControl)Field(form, "tabs")).SelectedIndex == 1 && !((TabPage)Field(form, "singleTab")).Enabled, "A parallel single comparison could be started.");
            PumpUntil(() => !Busy(batch), "six native Word comparisons");
            Finished(batch);
            IList records = Records(batch);
            Require(records.Count == 6, "Four files did not produce six pairs.");
            int[,] pairs = { {0,1}, {0,2}, {0,3}, {1,2}, {1,3}, {2,3} };
            for (int i = 0; i < 6; i++)
            {
                object record = records[i], pair = Property(record, "Pair");
                Require(State(record) == "Completed", "Pair failed: " + (Result(record) == null ? "no result" : Result(record).Error));
                int archive = (int)Property(pair, "ArchiveIndex"), revised = (int)Property(pair, "NewIndex");
                Require(archive == pairs[i,0] && revised == pairs[i,1], "Pair sequence or direction differs.");
                ComparisonResult result = Result(record);
                Require(result.NewWords == baseWords + revised && result.InsertedWords == revised - archive, "Wrong counts in pair " + (i + 1));
                Require(result.CrossoverPercent.HasValue && Math.Abs(result.CrossoverPercent.Value - 100.0 * (baseWords + archive) / (baseWords + revised)) < 0.051, "Wrong crossover in pair " + (i + 1));
                CheckOutput(record);
            }
            Require(((DataGridView)Field(batch, "resultsGrid")).Rows.Count == 6, "WF results grid must show six metric rows.");
            Require(File.ReadAllLines(Path.Combine(Output(batch), "batch-results.csv")).Length == 7, "WF CSV must contain header plus six records.");
            Require(((TabPage)Field(form, "singleTab")).Enabled, "Single comparison remained locked after batch.");
            Save(form, "batch-six-results.png");
            Console.WriteLine("PASS six native comparisons: requested order/direction, independently expected counts, crossover, yellow highlights, CSV, tab locking, unchanged sources.");
        });
    }
    private static void EpPair()
    {
        string input = Path.Combine(root, "ep-inputs");
        NewDoc(input, "01", "Description", LongEnglish, "Claims", "1. " + LongEnglish, "Revendications", "1. " + LongGerman);
        NewDoc(input, "02", "Description", LongEnglish, "Amber.", "Claims", "1. " + LongEnglish, "Cobalt.", "Revendications", "1. " + LongGerman);
        WithForm(input, "ep", true, delegate(Form form, object batch)
        {
            Click(batch, "startBatch");
            PumpUntil(() => !Busy(batch), "EP pair"); Finished(batch);
            object record = Records(batch)[0]; ComparisonResult result = Result(record);
            Require(State(record) == "Completed", "EP failed: " + (result == null ? "no result" : result.Error));
            Require(result.Sections != null && result.Sections.Length == 3, "EP scopes are missing.");
            Require(result.Sections[0].Scope == "Description" && result.Sections[0].InsertedWords == 1 && result.Sections[1].Scope == "Full specification" && result.Sections[1].InsertedWords == 2 && result.Sections[2].Scope == "Claims" && result.Sections[2].InsertedWords == 1, "EP insertion scopes are incorrect.");
            Require(((DataGridView)Field(batch, "resultsGrid")).Rows.Count == 3, "EP grid must have three metric rows.");
            Require(File.ReadAllLines(Path.Combine(Output(batch), "batch-results.csv")).Length == 4, "EP CSV must contain header plus three scope records.");
            Save(form, "batch-ep-results.png");
            Console.WriteLine("PASS EP batch: original claims selected, three scope rows in grid and CSV, expected insertion counts.");
        });
    }
    private static void ManualReview()
    {
        string input = Path.Combine(root, "short-inputs");
        NewDoc(input, "01", "Alpha beta gamma delta.");
        NewDoc(input, "02", "Alpha bright beta gamma delta.");
        WithForm(input, "manual", false, delegate(Form form, object batch)
        {
            Click(batch, "startBatch");
            string firstJob = (string)Field(batch, "jobDirectory");
            Require(!JsonFile.Read<ComparisonRequest>(Path.Combine(firstJob, "request.json")).MatchingLanguageConfirmed, "Batch silently bypassed language check.");
            PumpUntil(() => Busy(batch) && Field(batch, "worker") == null && ((Button)Field(batch, "languageChecked")).Visible, "manual review pause");
            Require(State(Records(batch)[0]) == "Review required", "Uncertain language pair did not pause for review.");
            Require(!Directory.Exists(firstJob) && Directory.GetFiles(Output(batch), "*.docx", SearchOption.AllDirectories).Length == 0, "Stopped pair retained temp content or published comparison.");
            Save(form, "batch-language-review.png");
            Click(batch, "languageChecked");
            string retryJob = (string)Field(batch, "jobDirectory");
            Require(JsonFile.Read<ComparisonRequest>(Path.Combine(retryJob, "request.json")).MatchingLanguageConfirmed, "Explicit manual confirmation did not reach worker.");
            PumpUntil(() => !Busy(batch), "manual retry"); Finished(batch);
            ComparisonResult result = Result(Records(batch)[0]);
            Require(State(Records(batch)[0]) == "Completed" && result.NewWords == 5 && result.InsertedWords == 1 && result.CrossoverPercent == 80.0, "Manual retry did not produce 80.0% / 5 / 1.");
            Require(result.Warnings.Any(w => w.Contains("Source language was confirmed manually")), "Manual review note missing.");
            Console.WriteLine("PASS manual review: batch stops, no output, explicit current-pair retry, expected metrics.");
        });
    }
    private static void FailedPairContinues()
    {
        string input = Path.Combine(root, "mismatch-inputs");
        NewDoc(input, "01", LongEnglish);
        NewDoc(input, "02", LongGerman);
        NewDoc(input, "03", LongEnglish);
        File.WriteAllText(Path.Combine(input, "04-corrupt.docx"), "This is an invalid ZIP document.");
        WithForm(input, "mismatch", false, delegate(Form form, object batch)
        {
            Click(batch, "startBatch"); PumpUntil(() => !Busy(batch), "mismatch and valid pair"); Finished(batch);
            IList records = Records(batch);
            Require(records.Count == 6 && State(records[1]) == "Completed" && records.Cast<object>().Count(r => State(r) == "Failed") == 5, "Known mismatches or corrupt input did not fail individually while valid pair continued.");
            Require(Result(records[0]).Error.Contains("Source language mismatch") && Result(records[3]).Error.Contains("Source language mismatch"), "Mismatch error missing.");
            Require(Result(records[1]).CrossoverPercent == 100.0 && Directory.GetFiles(Output(batch), "*.docx", SearchOption.AllDirectories).Length == 1, "Invalid language pairs created outputs or valid pair lost.");
            string csv = File.ReadAllText(Path.Combine(Output(batch), "batch-results.csv"));
            Require(csv.Contains("Failed") && csv.Contains("Source language mismatch"), "Combined report hid failures.");
            Console.WriteLine("PASS language mismatch and corrupt inputs: five failed pairs recorded, other pair completes, only valid comparison published.");
        });
    }
    private static void CancelPartialBatch()
    {
        string input = Path.Combine(root, "four-inputs");
        WithForm(input, "cancel", false, delegate(Form form, object batch)
        {
            Click(batch, "startBatch");
            PumpUntil(() => Records(batch).Count > 0 && State(Records(batch)[0]) == "Completed" && Field(batch, "worker") != null, "first completed pair and next active pair");
            string completedPath = Result(Records(batch)[0]).OutputPath, completedHash = Hash(completedPath);
            Click(batch, "cancelBatch"); PumpUntil(() => !Busy(batch), "batch cancellation"); Finished(batch);
            Require(Hash(completedPath) == completedHash, "Cancellation altered a completed comparison.");
            Require(Records(batch).Cast<object>().Any(r => State(r) == "Not run") && !Records(batch).Cast<object>().Any(r => State(r) == "Pending" || State(r) == "Running"), "Cancellation left pending or running records.");
            Require(Directory.GetFiles(Output(batch), "*.docx", SearchOption.AllDirectories).Length == Records(batch).Cast<object>().Count(r => State(r) == "Completed"), "Cancellation published an unrecorded completed output.");
            Console.WriteLine("PASS cancellation: private worker stops, completed document preserved, remaining pairs Not run, report saved.");
        });
    }
    private static void SkipAndCancelReview()
    {
        string input = Path.Combine(root, "skip-inputs");
        NewDoc(input, "01", "Alpha beta."); NewDoc(input, "02", "Alpha gamma."); NewDoc(input, "03", "Alpha delta.");
        WithForm(input, "skip", false, delegate(Form form, object batch)
        {
            Click(batch, "startBatch");
            PumpUntil(() => Field(batch, "worker") == null && ((Button)Field(batch, "skipPair")).Visible, "first review before skip");
            Click(batch, "skipPair");
            PumpUntil(() => Field(batch, "worker") == null && ((Button)Field(batch, "skipPair")).Visible && (int)Field(batch, "currentIndex") == 1, "next review after skip");
            Require(State(Records(batch)[0]) == "Failed", "Skip pair did not record incomplete failure.");
            Require((bool)batch.GetType().GetMethod("CancelAndClose").Invoke(batch, null), "Paused batch could not cancel.");
            Finished(batch);
            Require(!Records(batch).Cast<object>().Any(r => State(r) == "Pending" || State(r) == "Review required" || State(r) == "Running"), "Paused cancellation left unresolved states.");
            Require(Directory.GetFiles(Output(batch), "*.docx", SearchOption.AllDirectories).Length == 0, "Skipped or unconfirmed pair created output.");
            Console.WriteLine("PASS skip and paused cancellation: no silent language bypass, no comparison output, all statuses resolved.");
        });
    }
    [STAThread]
    private static int Main(string[] args)
    {
        try
        {
            if (args.Length != 2) throw new ArgumentException("Usage: BatchUiIntegration.exe <DocumentCompare.exe> <new local artifact folder>");
            app = Assembly.LoadFrom(Path.GetFullPath(args[0])); root = Path.GetFullPath(args[1]);
            // Builds use a unique staging filename before publication. Resolve the
            // resulting internal assembly name to the exact shipped file under test.
            AppDomain.CurrentDomain.AssemblyResolve += delegate(object sender, ResolveEventArgs resolve)
            { return new AssemblyName(resolve.Name).Name == app.GetName().Name ? app : null; };
            Require(!Directory.Exists(root), "Use a new artifact folder."); Directory.CreateDirectory(root);
            Application.EnableVisualStyles(); Application.SetCompatibleTextRenderingDefault(false);
            SixPairs(); EpPair(); ManualReview(); FailedPairContinues(); CancelPartialBatch(); SkipAndCancelReview();
            Console.WriteLine("PASS " + checks + " batch UI integration checks; app=" + app.GetName().Version + "; artifacts=" + root);
            return 0;
        }
        catch (Exception error) { Console.Error.WriteLine("FAIL " + error); return 1; }
    }
}
