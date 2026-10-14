using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Reflection;
using System.Security;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using System.Web.Script.Serialization;
using System.Windows.Forms;
using System.Xml.Linq;
using LocalDocumentCompare;

// Synthetic integration for an explicitly approved location. Never creates a
// share, mapping, or server. Run separately from other Word automation tests.
internal static class SharedDriveIntegration
{
    private const string English = "The present invention relates to a method for controlling a motor. The apparatus comprises a housing, a controller and a sensor. The controller receives the signal from the sensor and adjusts the speed of the motor. This arrangement improves the operation of the device. The following description explains how the components work together and how the operator can use the apparatus under different conditions. Each component performs a specific function within the system, and the controller records the result before the next operation begins.";
    private static string release, shared, local, alternate;
    private static Assembly app;
    private static int checks;
    private static readonly HashSet<string> observedJobs = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
    private static void Require(bool value, string message)
    { if (!value) throw new InvalidOperationException(message); checks++; }
    private static string Hash(string file)
    {
        using (SHA256 hash = SHA256.Create()) using (Stream stream = File.OpenRead(file))
            return BitConverter.ToString(hash.ComputeHash(stream));
    }
    private static object Field(object obj, string name)
    { return obj.GetType().GetField(name, BindingFlags.NonPublic | BindingFlags.Instance).GetValue(obj); }
    private static object Property(object obj, string name)
    { return obj.GetType().GetProperty(name).GetValue(obj, null); }
    private static void Add(ZipArchive zip, string name, string text)
    { using (var writer = new StreamWriter(zip.CreateEntry(name).Open(), new UTF8Encoding(false))) writer.Write(text); }
    private static string NewDoc(string folder, string name, params string[] paragraphs)
    {
        Directory.CreateDirectory(folder);
        string path = Path.Combine(folder, name + ".docx");
        using (var zip = ZipFile.Open(path, ZipArchiveMode.Create))
        {
            Add(zip, "[Content_Types].xml", "<Types xmlns='http://schemas.openxmlformats.org/package/2006/content-types'><Default Extension='rels' ContentType='application/vnd.openxmlformats-package.relationships+xml'/><Default Extension='xml' ContentType='application/xml'/><Override PartName='/word/document.xml' ContentType='application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml'/></Types>");
            Add(zip, "_rels/.rels", "<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'><Relationship Id='rId1' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument' Target='word/document.xml'/></Relationships>");
            string body = String.Join("", paragraphs.Select(p => "<w:p><w:r><w:t xml:space='preserve'>" + SecurityElement.Escape(p) + "</w:t></w:r></w:p>"));
            Add(zip, "word/document.xml", "<w:document xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'><w:body>" + body + "<w:sectPr/></w:body></w:document>");
        }
        return path;
    }
    private static bool Yellow(string path)
    {
        XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        using (var zip = ZipFile.OpenRead(path)) using (Stream stream = zip.GetEntry("word/document.xml").Open())
            return XDocument.Load(stream).Descendants(w + "ins").Descendants(w + "r")
                .Where(r => r.Element(w + "t") != null).Descendants(w + "highlight").Any(e => (string)e.Attribute(w + "val") == "yellow");
    }
    private static void OwnedWordClosed(string job)
    {
        string record = Path.Combine(job, "word.process");
        if (!File.Exists(record)) return;
        var identity = new JavaScriptSerializer().Deserialize<Dictionary<string, object>>(File.ReadAllText(record));
        int pid = Convert.ToInt32(identity["Id"]);
        try { using (Process p = Process.GetProcessById(pid)) Require(p.HasExited, "Private Word remains active after completion."); }
        catch (ArgumentException) { checks++; }
    }
    private static ComparisonResult Worker(ComparisonRequest request, string label)
    {
        string job = Path.Combine(local, label);
        Directory.CreateDirectory(job); request.WorkDirectory = job;
        string input = Path.Combine(job, "request.json"), output = Path.Combine(job, "result.json");
        JsonFile.Write(input, request);
        ProcessStartInfo start = new ProcessStartInfo(release, "--worker \"" + input + "\" \"" + output + "\"")
            { UseShellExecute = false, CreateNoWindow = true, WorkingDirectory = Path.GetDirectoryName(release) };
        using (Process worker = Process.Start(start))
        {
            if (!worker.WaitForExit(120000)) throw new InvalidOperationException("Private worker timed out; its exact process records are retained in " + job);
        }
        Require(File.Exists(output), "Worker result missing."); OwnedWordClosed(job);
        return JsonFile.Read<ComparisonResult>(output);
    }
    private static void Single()
    {
        string input = Path.Combine(shared, "single"), outputRoot = Path.Combine(shared, "results");
        Directory.CreateDirectory(outputRoot);
        string archive = NewDoc(input, "Archive", "Alpha beta gamma delta."), revised = NewDoc(input, "New", "Alpha bright beta gamma delta.");
        string archiveHash = Hash(archive), newHash = Hash(revised), output = Path.Combine(outputRoot, "New vs Archive.docx");
        string selectedArchive = archive, selectedOutput = output;
        if (alternate != null)
        {
            selectedArchive = Path.Combine(alternate, "single", "Archive.docx");
            selectedOutput = Path.Combine(alternate, "results", "New vs Archive.docx");
            Require(Hash(selectedArchive) == archiveHash && Hash(Path.Combine(alternate, "single", "New.docx")) == newHash, "Mapped and alternate paths must identify identical synthetic sources.");
        }
        var request = new ComparisonRequest { JobType = "WF", ArchivePath = selectedArchive, NewPath = revised, OutputPath = selectedOutput, MatchingLanguageConfirmed = true };
        ComparisonResult result = Worker(request, "single");
        Require(result.Success, "Single comparison failed: " + result.Error);
        Require(result.NewWords == 5 && result.InsertedWords == 1 && result.CrossoverPercent == 80.0, "Single shared comparison expected 5 / 1 / 80.0%.");
        Require(File.Exists(result.SummaryPath) && Yellow(output), "Shared output/summary or yellow insertion missing.");
        Require(Hash(Path.Combine(local, "single", "archive-source.docx")) == archiveHash && Hash(Path.Combine(local, "single", "new-source.docx")) == newHash, "Working copies must be exact source copies in the private local work folder.");
        Require(Hash(archive) == archiveHash && Hash(revised) == newHash, "Shared sources changed.");
        string comparisonHash = Hash(output), summaryHash = Hash(result.SummaryPath);
        ComparisonResult collision = Worker(request, "no-overwrite");
        Require(!collision.Success && collision.Error.Contains("already exists"), "Existing shared outputs must be rejected.");
        Require(Hash(output) == comparisonHash && Hash(result.SummaryPath) == summaryHash, "Existing outputs changed after collision.");
        MethodInfo validate = app.GetType("LocalDocumentCompare.ComparisonEngine").GetMethod("Validate", BindingFlags.Static | BindingFlags.NonPublic);
        request.OutputPath = Path.Combine(outputRoot, "locality-check.docx"); request.WorkDirectory = shared;
        bool rejected = false;
        try { validate.Invoke(null, new object[] { request }); }
        catch (TargetInvocationException e) { rejected = e.InnerException.Message.Contains("Temporary working files"); }
        // This explicit test only promises rejection for actual network roots;
        // a local fallback run still exercises read/write/output behavior.
        if (shared.StartsWith(@"\\") || new DriveInfo(Path.GetPathRoot(shared)).DriveType == DriveType.Network)
            Require(rejected, "Shared sources/results must not allow shared temporary work.");
        if (alternate != null)
        {
            request.WorkDirectory = alternate; rejected = false;
            try { validate.Invoke(null, new object[] { request }); }
            catch (TargetInvocationException e) { rejected = e.InnerException.Message.Contains("Temporary working files"); }
            Require(rejected, "Alternate shared path must not allow shared temporary work.");
            Require(Hash(Path.Combine(alternate, "results", "New vs Archive.docx")) == comparisonHash, "UNC and mapped saved comparison differ.");
        }
        Console.WriteLine("PASS single shared source/read/write: 5 words, 1 insertion, 80.0%, yellow markup, unchanged sources, exact local copies and no overwrite.");
    }
    private static void ObserveJob(object batch)
    {
        string job = (string)Field(batch, "jobDirectory");
        if (String.IsNullOrEmpty(job) || !observedJobs.Add(job)) return;
        string boundary = Path.GetFullPath(Path.Combine(Path.GetTempPath(), "LocalDocumentCompare")).TrimEnd('\\') + "\\";
        Require(Path.GetFullPath(job).StartsWith(boundary, StringComparison.OrdinalIgnoreCase), "Batch temporary work escaped the app's local temporary root.");
        Require(!job.StartsWith(@"\\") && new DriveInfo(Path.GetPathRoot(job)).DriveType != DriveType.Network, "Batch temporary files are on a shared location.");
    }
    private static void Pump(object batch, Func<bool> done, string description)
    {
        DateTime deadline = DateTime.UtcNow.AddMinutes(12);
        while (!done())
        {
            ObserveJob(batch); Application.DoEvents();
            if ((bool)Field(batch, "reportBlocked"))
                throw new InvalidOperationException("Shared CSV save/replacement blocked: " + ((TextBox)Field(batch, "details")).Text);
            if ((bool)Field(batch, "reviewPaused"))
                throw new InvalidOperationException("Controlled same-language batch unexpectedly requires review: " + ((TextBox)Field(batch, "details")).Text);
            if (DateTime.UtcNow > deadline) throw new InvalidOperationException("Timed out: " + description);
            Thread.Sleep(20);
        }
        Application.DoEvents();
    }
    private static void Batch()
    {
        string input = Path.Combine(shared, "batch-input"), output = Path.Combine(shared, "batch-output");
        string text = String.Join(" ", Enumerable.Repeat(English, 8));
        NewDoc(input, "01", text); NewDoc(input, "02", text, "Amber.");
        NewDoc(input, "03", text, "Amber cobalt."); NewDoc(input, "04", text, "Amber cobalt jade.");
        File.WriteAllText(Path.Combine(input, "ignored.txt"), "Synthetic non-document");
        File.WriteAllText(Path.Combine(input, "~$ignored.docx"), "Synthetic Word lock");
        NewDoc(Path.Combine(input, "nested"), "99", "Synthetic nested document");
        Dictionary<string, string> hashes = Directory.GetFiles(input, "*.docx").ToDictionary(p => p, p => Hash(p));
        List<BatchInput> loaded = BatchPlan.LoadFolder(input);
        Require(loaded.Count == 4 && loaded.Select(p => Path.GetFileName(p.Path)).SequenceEqual(new[] { "01.docx", "02.docx", "03.docx", "04.docx" }), "Shared batch folder ordering/filtering differs.");
        if (alternate != null)
        {
            List<BatchInput> alternateLoaded = BatchPlan.LoadFolder(Path.Combine(alternate, "batch-input"));
            Require(alternateLoaded.Count == 4 && alternateLoaded.Select(p => Path.GetFileName(p.Path)).SequenceEqual(loaded.Select(p => Path.GetFileName(p.Path))), "Alternate shared batch path ordering/filtering differs.");
            Require(alternateLoaded.Select(p => Hash(p.Path)).SequenceEqual(loaded.Select(p => Hash(p.Path))), "Mapped and alternate batch paths differ.");
        }
        int baseline = DocxContent.Analyze(loaded[0].Path).Words;
        using (Form form = (Form)Activator.CreateInstance(app.GetType("LocalDocumentCompare.MainForm", true)))
        {
            object batch = Field(form, "batch");
            string lastSnapshot = null;
            var diagnostic = new System.Threading.Timer(delegate(object ignored)
            {
                try
                {
                    int index = (int)Field(batch, "currentIndex");
                    string status = "not started";
                    IList list = (IList)Field(batch, "records");
                    if (index >= 0 && index < list.Count) status = (string)Property(list[index], "State");
                    string snapshot = "pair=" + (index + 1) + " status=" + status + " busy=" + Property(batch, "IsBusy") +
                        " worker=" + (Field(batch, "worker") != null) + " localJob=" + (Field(batch, "jobDirectory") != null) +
                        " reportBlocked=" + Field(batch, "reportBlocked") + " stopBlocked=" + Field(batch, "stopBlocked");
                    if (snapshot != lastSnapshot) { lastSnapshot = snapshot; Console.WriteLine("PROGRESS " + DateTime.UtcNow.ToString("o") + " " + snapshot); }
                }
                catch { }
            }, null, 1000, 3000);
            try
            {
                form.Opacity = 0; form.ShowInTaskbar = false; form.Show();
                ((TabControl)Field(form, "tabs")).SelectedIndex = 1;
                ((TextBox)Field(batch, "folderPath")).Text = input;
                ((TextBox)Field(batch, "outputPath")).Text = output;
                ((ComboBox)Field(batch, "jobType")).SelectedIndex = 0;
                ((Button)Field(batch, "loadFiles")).PerformClick(); Application.DoEvents();
                Require(((DataGridView)Field(batch, "fileGrid")).Rows.Count == 4, "UI failed to load the shared batch folder.");
                ((Button)Field(batch, "startBatch")).PerformClick(); Application.DoEvents();
                Require((bool)Property(batch, "IsBusy"), "Shared batch failed to start.");
                Pump(batch, () => !(bool)Property(batch, "IsBusy"), "six shared-drive Word pairs");
                BatchPlan plan = (BatchPlan)Field(batch, "plan");
                var records = ((IList)Field(batch, "records")).Cast<BatchRecord>().ToList();
                Require(records.Count == 6, "Shared batch does not contain six pairs.");
                int[,] order = { {0,1}, {0,2}, {0,3}, {1,2}, {1,3}, {2,3} };
                for (int i = 0; i < records.Count; i++)
                {
                    BatchRecord record = records[i]; ComparisonResult result = record.Result;
                    Require(record.Pair.ArchiveIndex == order[i,0] && record.Pair.NewIndex == order[i,1], "Shared pair order differs.");
                    Require(record.State == "Completed" && result != null && result.Success, "Shared batch pair failed: " + (result == null ? "no result" : result.Error));
                    Require(result.NewWords == baseline + order[i,1] && result.InsertedWords == order[i,1] - order[i,0], "Shared batch word counts differ.");
                    Require(result.CrossoverPercent.HasValue && Math.Abs(result.CrossoverPercent.Value - 100.0 * (baseline + order[i,0]) / (baseline + order[i,1])) < 0.051, "Shared batch crossover differs.");
                    Require(File.Exists(result.OutputPath) && File.Exists(result.SummaryPath) && Yellow(result.OutputPath), "A completed shared pair lacks yellow Word comparison or summary.");
                }
                Require(File.ReadAllLines(plan.ReportPath).Length == 7, "Shared CSV must have six records and one header.");
                Require(Field(batch, "jobDirectory") == null && observedJobs.Count == 6, "Shared batch retained a job or did not create a local job per pair.");
                Require(observedJobs.All(job => !Directory.Exists(job)), "Completed batch temporary work was not cleaned up.");
                Require(!Directory.GetFiles(plan.OutputDirectory, ".r*.tmp").Any(), "Shared CSV replacement left staging files.");
                Console.WriteLine("PROGRESS six native Word pairs completed; check owned shared report lock/retry.");
                string reportHash = Hash(plan.ReportPath);
                bool locked = false;
                Exception lockFailure = null;
                Thread save = null;
                using (var held = new FileStream(plan.ReportPath, FileMode.Open, FileAccess.Read, FileShare.None))
                {
                    save = new Thread(delegate() { try { BatchReport.Save(plan, records); } catch (Exception e) { lockFailure = e; } });
                    save.IsBackground = true; save.Start();
                    // Some SMB providers delay a replacement attempt while another
                    // handle denies sharing. Release our deliberate lock after 20s
                    // so the test cannot deadlock its own synchronous save call.
                    locked = save.Join(20000) && lockFailure is IOException;
                }
                Require(save.Join(45000), "Shared report save remained blocked after the deliberate lock was released.");
                Require(locked && Hash(plan.ReportPath) == reportHash, "Locked shared CSV was overwritten or lock not honoured.");
                BatchReport.Save(plan, records);
                Require(Hash(plan.ReportPath) == reportHash && !Directory.GetFiles(plan.OutputDirectory, ".r*.tmp").Any(), "Retry saving shared report changed rows or retained staging files.");
                foreach (var source in hashes) Require(Hash(source.Key) == source.Value, "A shared batch source changed.");
                CleanupBoundary(batch, input);
                Console.WriteLine("PASS shared batch: six native Word pairs in requested order, counts/crossover/yellow output, report replacements and lock/retry, unchanged sources, six local jobs cleaned and out-of-boundary cleanup refused.");
            }
            finally
            {
                diagnostic.Dispose();
                if ((bool)Property(batch, "IsBusy")) batch.GetType().GetMethod("CancelAndClose").Invoke(batch, null);
                form.Hide();
            }
        }
    }
    private static void CleanupBoundary(object batch, string sharedInput)
    {
        FieldInfo job = batch.GetType().GetField("jobDirectory", BindingFlags.NonPublic | BindingFlags.Instance);
        object previous = job.GetValue(batch);
        try
        {
            job.SetValue(batch, sharedInput);
            object retained = batch.GetType().GetMethod("CleanupJob", BindingFlags.NonPublic | BindingFlags.Instance).Invoke(batch, null);
            Require((string)retained == sharedInput && File.Exists(Path.Combine(sharedInput, "01.docx")), "Cleanup reached outside the app's owned local temporary root.");
        }
        finally { job.SetValue(batch, previous); }
    }
    private static int Run()
    {
        Application.EnableVisualStyles(); Application.SetCompatibleTextRenderingDefault(false);
        Single(); Batch();
        Console.WriteLine("PASS " + checks + " shared-location integration checks; version=" + app.GetName().Version + "; root=" + shared);
        return 0;
    }
    [STAThread]
    private static int Main(string[] args)
    {
        try
        {
            if (args.Length < 3 || args.Length > 4) throw new ArgumentException("Usage: SharedDriveIntegration.exe <release exe> <new owned shared folder> <new owned private local folder> [alternate path to same owned shared folder]");
            release = Path.GetFullPath(args[0]); shared = Path.GetFullPath(args[1]); local = Path.GetFullPath(args[2]);
            alternate = args.Length == 4 ? Path.GetFullPath(args[3]) : null;
            app = Assembly.LoadFrom(release);
            AppDomain.CurrentDomain.AssemblyResolve += delegate(object sender, ResolveEventArgs e)
            { return new AssemblyName(e.Name).Name == app.GetName().Name ? app : null; };
            return Run();
        }
        catch (Exception e) { Console.Error.WriteLine("FAIL " + e); return 1; }
    }
}
