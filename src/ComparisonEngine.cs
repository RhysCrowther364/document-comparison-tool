using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;

namespace LocalDocumentCompare
{
    public static class ComparisonEngine
    {
        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr hwnd, out uint processId);

        public static ComparisonResult Run(ComparisonRequest request, Action<string> progress)
        {
            var warnings = new List<string>();
            var result = new ComparisonResult { Warnings = new string[0] };
            dynamic word = null;
            dynamic archive = null;
            dynamic revised = null;
            dynamic comparison = null;
            bool ownsWord = false;
            int ownedWordPid = 0;
            long ownedWordStart = 0;
            bool optionsCaptured = false;
            bool oldLinks = false, oldFieldsAtPrint = false, oldLinksAtPrint = false, oldNormalPrompt = false;
            string stage = "validating files";
            bool outputCreated = false, summaryCreated = false;
            string outputPath = null, summaryPath = null;
            try
            {
                Validate(request);
                bool ep = String.Equals(request.JobType, "EP", StringComparison.OrdinalIgnoreCase);
                result.JobType = ep ? "EP" : "WF";
                outputPath = Path.GetFullPath(request.OutputPath);
                summaryPath = Path.ChangeExtension(outputPath, ".txt");
                Notify(progress, "Preparing local working copies...");
                string archiveCopy = CopyInput(request.ArchivePath, "archive", request.WorkDirectory);
                string revisedCopy = CopyInput(request.NewPath, "new", request.WorkDirectory);
                foreach (string path in new[] { archiveCopy, revisedCopy })
                    if (IsOoxml(path) && DocxContent.DetachExternalRelationships(path))
                        warnings.Add("External links were detached from a temporary input copy; available cached text and images were retained.");

                stage = "checking document languages";
                Notify(progress, "Checking source languages locally...");
                CheckCancelled(request.WorkDirectory);
                CheckEarlyLanguages(request, result, archiveCopy, revisedCopy, ep, warnings);

                stage = "starting desktop Microsoft Word";
                Notify(progress, "Starting a separate local Word instance...");
                CheckCancelled(request.WorkDirectory);
                var existingIds = new HashSet<int>(Process.GetProcessesByName("WINWORD").Select(p => { try { return p.Id; } finally { p.Dispose(); } }));
                Type wordType = Type.GetTypeFromProgID("Word.Application");
                if (wordType == null) throw new UserError("Desktop Microsoft Word is required. Install and activate Word, then try again.");
                word = Activator.CreateInstance(wordType);
                uint wordPid = 0;
                // Word exposes Hwnd on Window, not Application. A private blank document
                // gives us a window handle before any input document is opened.
                dynamic bootstrap = null;
                dynamic bootstrapWindow = null;
                dynamic bootstrapDocuments = null;
                dynamic bootstrapWindows = null;
                try
                {
                    stage = "accessing Word's document collection";
                    bootstrapDocuments = word.Documents;
                    stage = "creating the private Word window";
                    bootstrap = bootstrapDocuments.Add();
                    stage = "accessing the private Word window";
                    bootstrapWindows = bootstrap.Windows;
                    bootstrapWindow = bootstrapWindows.Item(1);
                    stage = "identifying the private Word window";
                    GetWindowThreadProcessId(new IntPtr((int)bootstrapWindow.Hwnd), out wordPid);
                }
                finally
                {
                    Release((object)bootstrapWindow);
                    Release((object)bootstrapWindows);
                    CloseDocument(ref bootstrap);
                    Release((object)bootstrapDocuments);
                }
                if (wordPid == 0 || existingIds.Contains((int)wordPid))
                    throw new UserError("A separate Word process could not be verified. Close this tool, check Word is working, and try again. Your existing Word windows were left alone.");
                ownsWord = true;
                ownedWordPid = (int)wordPid;
                using (Process ownedProcess = Process.GetProcessById(ownedWordPid)) ownedWordStart = ownedProcess.StartTime.ToUniversalTime().Ticks;
                WriteProcessIdentity(request.WorkDirectory, "word", (int)wordPid);
                stage = "configuring the private Word session";
                CheckCancelled(request.WorkDirectory);
                word.Visible = false;
                word.DisplayAlerts = 0; // wdAlertsNone
                word.AutomationSecurity = 3; // msoAutomationSecurityForceDisable
                word.ScreenUpdating = false;
                dynamic options = word.Options;
                try
                {
                    oldLinks = options.UpdateLinksAtOpen;
                    oldFieldsAtPrint = options.UpdateFieldsAtPrint;
                    oldLinksAtPrint = options.UpdateLinksAtPrint;
                    oldNormalPrompt = options.SaveNormalPrompt;
                    optionsCaptured = true;
                    options.UpdateLinksAtOpen = false;
                    options.UpdateFieldsAtPrint = false;
                    options.UpdateLinksAtPrint = false;
                    options.SaveNormalPrompt = false;
                }
                finally { Release((object)options); }

                stage = "opening the archive";
                CheckCancelled(request.WorkDirectory);
                Notify(progress, request.OcrArchive ? "Recognising archive PDF text locally..." : "Opening archive in Word...");
                archive = OpenInput(word, archiveCopy, request.OcrArchive, "archive", request.WorkDirectory, warnings);
                stage = "preparing the archive's final text";
                PrepareFinalInput(archive, "archive", warnings);
                string archiveDocx = Path.Combine(request.WorkDirectory, "archive-final.docx");
                SaveDocx(archive, archiveDocx);
                if (!ep) ValidateConvertedPdf(archiveCopy, request.OcrArchive, archiveDocx, "archive");

                stage = "opening the new document";
                CheckCancelled(request.WorkDirectory);
                Notify(progress, request.OcrNew ? "Recognising new PDF text locally..." : "Opening new document in Word...");
                revised = OpenInput(word, revisedCopy, request.OcrNew, "new", request.WorkDirectory, warnings);
                stage = "preparing the new document's final text";
                PrepareFinalInput(revised, "new", warnings);
                string revisedDocx = Path.Combine(request.WorkDirectory, "new-final.docx");
                SaveDocx(revised, revisedDocx);
                if (!ep) ValidateConvertedPdf(revisedCopy, request.OcrNew, revisedDocx, "new");
                if (ep)
                {
                    stage = "identifying the EP description and claims";
                    Notify(progress, "Preparing EP description and original-language claims...");
                    CheckCancelled(request.WorkDirectory);
                    CloseDocument(ref archive);
                    CloseDocument(ref revised);
                    string scopedArchive = Path.Combine(request.WorkDirectory, "archive-ep.docx");
                    string scopedNew = Path.Combine(request.WorkDirectory, "new-ep.docx");
                    EpSections.Prepare(archiveDocx, scopedArchive, "archive");
                    EpSections.Prepare(revisedDocx, scopedNew, "new");
                    archiveDocx = scopedArchive;
                    revisedDocx = scopedNew;
                    ValidateConvertedPdf(archiveCopy, request.OcrArchive, archiveDocx, "archive", true);
                    ValidateConvertedPdf(revisedCopy, request.OcrNew, revisedDocx, "new", true);
                    stage = "checking EP source languages";
                    Notify(progress, "Checking description and original-claims languages locally...");
                    ValidateLanguages(request, result, ReadLanguagePairs(archiveDocx, revisedDocx, true, request.WorkDirectory, "final-language"), warnings);
                    archive = OpenInput(word, archiveDocx, false, "archive", request.WorkDirectory, warnings);
                    revised = OpenInput(word, revisedDocx, false, "new", request.WorkDirectory, warnings);
                }
                else
                {
                    stage = "checking document languages";
                    Notify(progress, "Checking source languages locally...");
                    ValidateLanguages(request, result, ReadLanguagePairs(archiveDocx, revisedDocx, false, request.WorkDirectory, "final-language"), warnings);
                }
                TextMetrics newText = ep ? AnalyzeEpScope(revisedDocx, request.WorkDirectory, "new", "FullSpecification") : DocxContent.Analyze(revisedDocx);
                if (newText.Words == 0) warnings.Add("The new document has no countable text. Crossover is undefined.");

                stage = "comparing the documents in Word";
                CheckCancelled(request.WorkDirectory);
                Notify(progress, "Comparing text in Word; large documents can take several minutes...");
                comparison = word.CompareDocuments(
                    OriginalDocument: archive,
                    RevisedDocument: revised,
                    Destination: 2, // wdCompareDestinationNew
                    Granularity: 1, // wdGranularityWordLevel
                    CompareFormatting: false,
                    CompareCaseChanges: true,
                    CompareWhitespace: true,
                    CompareTables: true,
                    CompareHeaders: true,
                    CompareFootnotes: true,
                    CompareTextboxes: true,
                    CompareFields: true,
                    CompareComments: false,
                    CompareMoves: true,
                    RevisedAuthor: "Comparison",
                    IgnoreAllComparisonWarnings: true);
                comparison.TrackRevisions = false;
                comparison.ShowRevisions = true;
                dynamic windows = comparison.Windows;
                try
                {
                    for (int i = 1; i <= (int)windows.Count; i++)
                    {
                        dynamic window = windows.Item(i);
                        dynamic view = null;
                        try
                        {
                            view = window.View;
                            view.ShowRevisionsAndComments = true;
                            view.RevisionsView = 0; // final with markup; saved revision data remains in the file
                            view.ShowInsertionsAndDeletions = true;
                            view.ShowFormatChanges = false;
                        }
                        finally { Release((object)view); Release((object)window); }
                    }
                }
                finally { Release((object)windows); }
                string comparedDocx = Path.Combine(request.WorkDirectory, "comparison.docx");
                SaveDocx(comparison, comparedDocx);
                CloseDocument(ref comparison);
                if (DocxContent.DetachExternalRelationships(comparedDocx))
                    warnings.Add("External links were detached from the exported comparison; cached display content was retained.");
                stage = "highlighting and measuring the comparison";
                Notify(progress, "Highlighting insertions and calculating the summary...");
                TextMetrics comparedText = ep ? AnalyzeEpScope(comparedDocx, request.WorkDirectory, "compared", "FullSpecification") : DocxContent.Analyze(comparedDocx);
                DocxContent.HighlightInsertionsAndShowMarkup(comparedDocx);
                ComparisonSectionResult full = MeasureScope(ep ? "Full specification" : "Document", newText, comparedText, warnings);
                result.NewWords = full.NewWords;
                result.InsertedWords = full.InsertedWords;
                result.CrossoverPercent = full.CrossoverPercent;
                if (ep)
                {
                    var sections = new List<ComparisonSectionResult>();
                    foreach (string scope in new[] { "Description", "Claims" })
                    {
                        TextMetrics scopeNew = AnalyzeEpScope(revisedDocx, request.WorkDirectory, "new", scope);
                        TextMetrics scopeCompared = AnalyzeEpScope(comparedDocx, request.WorkDirectory, "compared", scope);
                        sections.Add(MeasureScope(scope, scopeNew, scopeCompared, warnings));
                    }
                    if (sections.Sum(s => s.NewWords) != full.NewWords || sections.Sum(s => s.InsertedWords) != full.InsertedWords)
                        throw new UserError("The EP section counts could not be reconciled with the full specification. Prepare copies containing a clearly labelled Description and original-language Claims section, then retry.");
                    result.Sections = new[] { sections[0], full, sections[1] };
                }
                if (IsPdf(archiveCopy) || IsPdf(revisedCopy))
                    warnings.Add("PDF conversion can change reading order, line breaks, tables and layout. Review the converted text and comparison before relying on the figures.");
                result.Warnings = warnings.Distinct(StringComparer.Ordinal).ToArray();
                result.OutputPath = outputPath;
                result.SummaryPath = summaryPath;
                stage = "saving the results";
                CheckCancelled(request.WorkDirectory);
                // CreateNew gives a final atomic no-overwrite check. Record ownership before copying
                // so disk-full and other mid-copy failures remove only this invocation's partial file.
                using (var destination = new FileStream(outputPath, FileMode.CreateNew, FileAccess.Write, FileShare.None))
                {
                    outputCreated = true;
                    using (var source = File.OpenRead(comparedDocx)) source.CopyTo(destination);
                }
                using (var stream = new FileStream(summaryPath, FileMode.CreateNew, FileAccess.Write, FileShare.None))
                {
                    summaryCreated = true;
                    using (var writer = new StreamWriter(stream, new UTF8Encoding(false)))
                        writer.Write(BuildSummary(request, result));
                }
                result.Success = true;
                Notify(progress, "Comparison and summary saved to the selected folder.");
            }
            catch (UserError ex)
            {
                result.Error = ex.Message;
            }
            catch (EpSectionException ex)
            {
                result.Error = ex.Message;
            }
            catch (Exception ex)
            {
                // Do not log Word's exception text: it can contain filenames or document content.
                result.Error = stage == "saving the results"
                    ? "Could not save the results. Check that the selected folder is available and writable, that any shared-drive or VPN connection is working, and that there is enough free space. Technical code: " + ex.HResult.ToString("X8", CultureInfo.InvariantCulture) + "."
                    : "Could not finish " + stage + ". Check that the files are readable, unprotected and supported, that any shared-drive connection is working, and that desktop Word can open them. Close any Word dialog and try again. Technical code: " + ex.HResult.ToString("X8", CultureInfo.InvariantCulture) + ".";
            }
            finally
            {
                if (ownsWord)
                {
                    CloseDocument(ref comparison);
                    CloseDocument(ref revised);
                    CloseDocument(ref archive);
                    try
                    {
                        if (optionsCaptured)
                        {
                            dynamic options = word.Options;
                            try
                            {
                                options.UpdateLinksAtOpen = oldLinks;
                                options.UpdateFieldsAtPrint = oldFieldsAtPrint;
                                options.UpdateLinksAtPrint = oldLinksAtPrint;
                                options.SaveNormalPrompt = oldNormalPrompt;
                            }
                            finally { Release((object)options); }
                        }
                    }
                    catch { }
                    try { word.Quit(SaveChanges: 0); } catch { }
                }
                Release((object)word);
                // A failed partial export belongs to this invocation only.
                if (!result.Success)
                {
                    if (outputCreated) TryDelete(outputPath);
                    if (summaryCreated) TryDelete(summaryPath);
                    result.OutputPath = null;
                    result.SummaryPath = null;
                    result.Warnings = warnings.Distinct(StringComparer.Ordinal).ToArray();
                }
                GC.Collect();
                GC.WaitForPendingFinalizers();
                if (ownsWord)
                {
                    try
                    {
                        using (Process ownedProcess = Process.GetProcessById(ownedWordPid))
                        {
                            if (ownedProcess.StartTime.ToUniversalTime().Ticks == ownedWordStart && !ownedProcess.WaitForExit(5000))
                            {
                                ownedProcess.Kill();
                                ownedProcess.WaitForExit(3000);
                            }
                        }
                    }
                    catch (ArgumentException) { } // Word has already exited.
                    catch { warnings.Add("The private Word process could not be confirmed closed. Close this tool's Word session before removing its temporary files."); }
                }
                result.Warnings = warnings.Distinct(StringComparer.Ordinal).ToArray();
            }
            return result;
        }

        private sealed class LanguagePair
        {
            public string Scope;
            public SourceLanguage Archive, New;
            public bool ArchiveHasText, NewHasText;
        }

        private static void CheckEarlyLanguages(ComparisonRequest request, ComparisonResult result, string archive, string revised, bool ep, List<string> warnings)
        {
            // PDF and legacy DOC need conversion first. Readable OOXML can be
            // checked before even starting Word; repeat the check on Word's
            // canonical final copies later for every supported input format.
            if (!IsOoxml(archive) || !IsOoxml(revised)) return;
            List<LanguagePair> pairs;
            try
            {
                if (ep)
                {
                    string scopedArchive = Path.Combine(request.WorkDirectory, "archive-language-preflight.docx");
                    string scopedNew = Path.Combine(request.WorkDirectory, "new-language-preflight.docx");
                    EpSections.Prepare(archive, scopedArchive, "archive");
                    EpSections.Prepare(revised, scopedNew, "new");
                    archive = scopedArchive;
                    revised = scopedNew;
                }
                pairs = ReadLanguagePairs(archive, revised, ep, request.WorkDirectory, "preflight-language");
            }
            catch (Exception)
            {
                // Finalization can resolve tracked headings or older OOXML.
                // Let the authoritative conversion/EP checks handle those cases.
                return;
            }
            ValidateLanguages(request, result, pairs, warnings);
        }

        private static List<LanguagePair> ReadLanguagePairs(string archive, string revised, bool ep, string workDirectory, string prefix)
        {
            var pairs = new List<LanguagePair>();
            foreach (string scope in ep ? new[] { "Description", "Claims" } : new[] { "Document" })
            {
                string archivePath = archive, newPath = revised;
                if (ep)
                {
                    archivePath = Path.Combine(workDirectory, prefix + "-archive-" + scope + ".docx");
                    newPath = Path.Combine(workDirectory, prefix + "-new-" + scope + ".docx");
                    EpSections.ExtractScope(archive, archivePath, scope);
                    EpSections.ExtractScope(revised, newPath, scope);
                }
                string archiveText = DocxContent.ReadFinalText(archivePath);
                string newText = DocxContent.ReadFinalText(newPath);
                pairs.Add(new LanguagePair
                {
                    Scope = scope, Archive = LanguageDetection.Detect(archiveText), New = LanguageDetection.Detect(newText),
                    ArchiveHasText = !String.IsNullOrWhiteSpace(archiveText), NewHasText = !String.IsNullOrWhiteSpace(newText)
                });
            }
            return pairs;
        }

        private static void ValidateLanguages(ComparisonRequest request, ComparisonResult result, List<LanguagePair> pairs, List<string> warnings)
        {
            // Check every matching EP scope before allowing any manual retry.
            // A manual check can resolve uncertainty, never a detected mismatch.
            foreach (LanguagePair pair in pairs)
                if (pair.ArchiveHasText && pair.NewHasText && pair.Archive.Status == LanguageStatus.Known && pair.New.Status == LanguageStatus.Known &&
                    !String.Equals(pair.Archive.Code, pair.New.Code, StringComparison.OrdinalIgnoreCase))
                {
                    result.LanguageMismatch = new SourceLanguageMismatch
                    {
                        Scope = pair.Scope, ArchiveLanguage = pair.Archive.DisplayName, NewLanguage = pair.New.DisplayName
                    };
                    throw new UserError("Source language mismatch in " + pair.Scope + ": archive is " + pair.Archive.DisplayName + "; new is " + pair.New.DisplayName + ". Select documents in the same source language. No comparison has been created.");
                }

            var review = new List<string>();
            foreach (LanguagePair pair in pairs)
            {
                // An empty WF side has no source language to compare; retain
                // the established empty-document/deletion-only behaviour.
                if (!pair.ArchiveHasText || !pair.NewHasText) continue;
                if (pair.Archive.Status != LanguageStatus.Known || pair.New.Status != LanguageStatus.Known)
                    review.Add(pair.Scope + " (archive: " + pair.Archive.DisplayName + "; new: " + pair.New.DisplayName + ")");
            }
            if (pairs.Count > 1)
            {
                if (pairs.Where(p => p.Archive.Status == LanguageStatus.Known).Select(p => p.Archive.Code).Distinct(StringComparer.OrdinalIgnoreCase).Count() > 1)
                    review.Add("the archive description and selected claims use different languages");
                if (pairs.Where(p => p.New.Status == LanguageStatus.Known).Select(p => p.New.Code).Distinct(StringComparer.OrdinalIgnoreCase).Count() > 1)
                    review.Add("the new description and selected claims use different languages");
            }
            if (review.Count == 0) return;
            if (!request.MatchingLanguageConfirmed)
            {
                result.LanguageReviewRequired = true;
                throw new UserError("A manual source-language check is required for " + String.Join("; ", review) + ". Automatic detection cannot reliably verify short, mixed-language or unclear text. Check that both files use the same language in each compared section and, for EP, that the original-language claims are selected. Then click Language checked — compare. No comparison has been created.");
            }
            warnings.Add("Source language was confirmed manually for " + String.Join("; ", review) + ". Review the selected text and OCR accuracy before relying on the figures.");
        }

        private static void Validate(ComparisonRequest request)
        {
            if (request == null) throw new UserError("No comparison request was provided.");
            if (!String.IsNullOrWhiteSpace(request.JobType) && !String.Equals(request.JobType, "WF", StringComparison.OrdinalIgnoreCase) && !String.Equals(request.JobType, "EP", StringComparison.OrdinalIgnoreCase))
                throw new UserError("Select WF or EP as the job type.");
            foreach (string path in new[] { request.ArchivePath, request.NewPath })
            {
                RequireSourcePath(path);
                if (!File.Exists(path)) throw new UserError("An input file could not be found or read. Check file access and any shared-drive connection, then select both files again.");
                string extension = Path.GetExtension(path).ToLowerInvariant();
                if (extension != ".doc" && extension != ".docx" && extension != ".pdf")
                    throw new UserError("Choose a PDF, DOC or DOCX input file.");
            }
            if (request.OcrArchive && !IsPdf(request.ArchivePath) || request.OcrNew && !IsPdf(request.NewPath))
                throw new UserError("Local OCR can only be selected for a PDF input.");
            RequireOutputPath(request.OutputPath);
            RequireLocalPath(request.WorkDirectory);
            if (!Directory.Exists(request.WorkDirectory)) throw new UserError("The local working folder is unavailable.");
            string output = Path.GetFullPath(request.OutputPath);
            if (!String.Equals(Path.GetExtension(output), ".docx", StringComparison.OrdinalIgnoreCase))
                throw new UserError("The comparison output must be a DOCX file.");
            if (!Directory.Exists(Path.GetDirectoryName(output))) throw new UserError("The output folder is unavailable. Check the folder path and any shared-drive connection.");
            if (File.Exists(output) || File.Exists(Path.ChangeExtension(output, ".txt")))
                throw new UserError("A comparison or summary with this name already exists. Choose a different output name or folder.");
            if (String.Equals(output, Path.GetFullPath(request.ArchivePath), StringComparison.OrdinalIgnoreCase) ||
                String.Equals(output, Path.GetFullPath(request.NewPath), StringComparison.OrdinalIgnoreCase))
                throw new UserError("The output must have a different path from both input files.");
        }

        private static string RequireFileSystemPath(string path)
        {
            // Accept full File Explorer paths, including UNC shares and mapped drives.
            // Reject drive-relative paths, web links and device namespaces.
            string normalized = (path ?? "").Replace('/', '\\');
            if (!Regex.IsMatch(normalized, @"^[A-Za-z]:\\") &&
                !Regex.IsMatch(normalized, @"^\\\\[^\\/:*?""<>|]+\\[^\\/:*?""<>|]+(?:\\|$)"))
                throw new UserError("Choose a full file or folder path from File Explorer, such as C:\\Folder or \\\\server\\share\\Folder.");
            if (normalized.StartsWith(@"\\.\", StringComparison.Ordinal))
                throw new UserError("Choose a regular file or folder path from File Explorer.");
            return Path.GetFullPath(normalized);
        }

        internal static void RequireSourcePath(string path)
        {
            RequireFileSystemPath(path);
        }

        internal static void RequireLocalPath(string path)
        {
            string full = RequireFileSystemPath(path);
            if (full.StartsWith(@"\\", StringComparison.Ordinal) || new DriveInfo(Path.GetPathRoot(full)).DriveType == DriveType.Network)
                throw new UserError("Temporary working files must stay on this PC, outside network locations. Source files and saved results can use shared drives.");
            if (IsKnownOneDrivePath(full))
                throw new UserError("Temporary working files must stay outside OneDrive. Source files and saved results can use OneDrive.");
        }

        internal static void RequireOutputPath(string path)
        {
            // Publishing to shared/synchronised storage is a user choice. Conversion,
            // OCR and comparison continue to use separately validated private local work.
            RequireFileSystemPath(path);
        }

        private static bool IsKnownOneDrivePath(string path)
        {
            string full = Path.GetFullPath(path);
            foreach (string name in new[] { "OneDrive", "OneDriveConsumer", "OneDriveCommercial" })
            {
                string cloudRoot = Environment.GetEnvironmentVariable(name);
                if (String.IsNullOrWhiteSpace(cloudRoot)) continue;
                cloudRoot = Path.GetFullPath(cloudRoot).TrimEnd(Path.DirectorySeparatorChar) + Path.DirectorySeparatorChar;
                if (full.StartsWith(cloudRoot, StringComparison.OrdinalIgnoreCase) ||
                    String.Equals(full.TrimEnd(Path.DirectorySeparatorChar), cloudRoot.TrimEnd(Path.DirectorySeparatorChar), StringComparison.OrdinalIgnoreCase))
                    return true;
            }
            return false;
        }

        private static string CopyInput(string source, string label, string workingDirectory)
        {
            string path = Path.Combine(workingDirectory, label + "-source" + Path.GetExtension(source).ToLowerInvariant());
            try { File.Copy(source, path, false); }
            catch (Exception ex)
            {
                if (!(ex is IOException) && !(ex is UnauthorizedAccessException)) throw;
                if (IsKnownOneDrivePath(source))
                    throw new UserError("The " + label + " file could not be read from OneDrive. In File Explorer, right-click it, choose Always keep on this device, and wait for the download to finish. Check your OneDrive connection and file permissions, then retry.");
                throw new UserError("The " + label + " file could not be copied to the local working folder. Check file permissions, any shared-drive or VPN connection, and free space on this PC, then retry.");
            }
            File.SetAttributes(path, FileAttributes.Normal);
            return path;
        }

        private static object OpenInput(dynamic word, string source, bool ocr, string label, string workingDirectory, List<string> warnings)
        {
            if (ocr)
            {
                string textPath = RunOcr(source, label, workingDirectory);
                string content = File.ReadAllText(textPath, Encoding.UTF8);
                if (!Regex.IsMatch(content, @"[\p{L}\p{N}]"))
                    throw new UserError("Local OCR did not recognise text in the " + label + " PDF. Review its scan quality and installed Windows OCR language, then try again.");
                string recognizedDocx = Path.Combine(workingDirectory, label + "-recognised.docx");
                bool layoutRestored = OcrLayout.TryCreateDocument(recognizedDocx, content, Path.ChangeExtension(textPath, ".layout.json"));
                warnings.Add("The " + label + " PDF was recognised with local Windows OCR. All pages were processed, including selectable-text pages. " +
                    (layoutRestored ? "Page sizes, text sizes, indents and spacing were reconstructed where possible. Fonts, complex columns, tables and graphics may differ; review against the PDF." :
                    "Page layout could not be reconstructed, so simple text formatting was used. Review spelling, reading order and layout against the PDF."));
                source = recognizedDocx;
            }
            // A random password prevents a hidden password prompt for encrypted inputs.
            string impossiblePassword = Guid.NewGuid().ToString("N");
            dynamic collection = word.Documents;
            try
            {
                return collection.Open(FileName: source, ConfirmConversions: false, ReadOnly: false,
                    AddToRecentFiles: false, PasswordDocument: impossiblePassword, PasswordTemplate: impossiblePassword,
                    Revert: false, WritePasswordDocument: impossiblePassword, WritePasswordTemplate: impossiblePassword,
                    Visible: false, OpenAndRepair: false, NoEncodingDialog: true);
            }
            finally { Release((object)collection); }
        }

        private static void PrepareFinalInput(dynamic document, string label, List<string> warnings)
        {
            if ((int)document.ProtectionType != -1)
                throw new UserError("The " + label + " document is protected. Use an authorised unprotected local copy.");
            document.TrackRevisions = false;
            bool acceptedRevisions = false;
            // Prevent cached fields in legacy DOC inputs updating during save/compare.
            dynamic stories = document.StoryRanges;
            try
            {
                for (int storyType = 1; storyType <= 17; storyType++)
                {
                    dynamic range = null;
                    try { range = stories.Item(storyType); }
                    catch (COMException) { continue; }
                    while (range != null)
                    {
                        dynamic next = null;
                        dynamic fields = null;
                        dynamic revisions = null;
                        try
                        {
                            revisions = range.Revisions;
                            if ((int)revisions.Count > 0)
                            {
                                revisions.AcceptAll();
                                acceptedRevisions = true;
                            }
                            fields = range.Fields;
                            if ((int)fields.Count > 0) fields.Locked = true;
                            next = range.NextStoryRange;
                        }
                        finally { Release((object)revisions); Release((object)fields); Release((object)range); }
                        range = next;
                    }
                }
            }
            finally { Release((object)stories); }
            if (acceptedRevisions)
                warnings.Add("Existing tracked changes in the " + label + " input were accepted on its temporary copy to compare final text. The original file was unchanged.");
        }

        private static void ValidateConvertedPdf(string original, bool usedOcr, string docx, string label, bool epScoped = false)
        {
            if (!IsPdf(original) || usedOcr) return;
            TextMetrics text = DocxContent.Analyze(docx);
            if (text.Words == 0)
                throw new UserError("The " + label + " PDF did not produce readable text in Word. Enable local OCR for that input and compare again.");
            if (text.HasImages || (!epScoped && PdfMayContainImages(original)))
                throw new UserError("The " + label + " PDF contains images that may hold scanned text. Enable local OCR for that input so scanned pages are included. This check is conservative and can also flag logos or diagrams.");
        }

        private static bool PdfMayContainImages(string path)
        {
            // Supplement Word's conversion output. PDF object streams may hide this marker;
            // converted drawing detection above is the primary safeguard.
            using (var reader = new StreamReader(path, Encoding.GetEncoding(28591), false))
            {
                char[] buffer = new char[65536];
                string tail = "";
                int read;
                var marker = new Regex(@"/Subtype\s*/Image\b", RegexOptions.CultureInvariant);
                while ((read = reader.Read(buffer, 0, buffer.Length)) > 0)
                {
                    string chunk = tail + new String(buffer, 0, read);
                    if (marker.IsMatch(chunk)) return true;
                    tail = chunk.Substring(Math.Max(0, chunk.Length - 128));
                }
            }
            return false;
        }

        private static string RunOcr(string input, string label, string workingDirectory)
        {
            string script = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "scripts", "Convert-PdfLocal.ps1");
            if (!File.Exists(script)) throw new UserError("The local OCR component is missing. Reinstall the complete application folder.");
            string output = Path.Combine(workingDirectory, label + "-ocr.txt");
            string status = Path.Combine(workingDirectory, label + "-ocr-status.json");
            string layout = Path.ChangeExtension(output, ".layout.json");
            string powershell = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.System), "WindowsPowerShell", "v1.0", "powershell.exe");
            CheckCancelled(workingDirectory);
            var start = new ProcessStartInfo
            {
                FileName = powershell,
                Arguments = "-NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass -File " + QuoteArgument(script) +
                    " -InputPath " + QuoteArgument(input) + " -OutputPath " + QuoteArgument(output) + " -StatusPath " + QuoteArgument(status) + " -LayoutPath " + QuoteArgument(layout),
                UseShellExecute = false,
                CreateNoWindow = true,
                WindowStyle = ProcessWindowStyle.Hidden,
                RedirectStandardOutput = true,
                RedirectStandardError = true
            };
            using (Process process = Process.Start(start))
            {
                try { WriteProcessIdentity(workingDirectory, "ocr", process.Id); }
                catch { try { process.Kill(); } catch { } throw; }
                if (File.Exists(Path.Combine(workingDirectory, "cancel.flag")))
                {
                    try { process.Kill(); } catch { }
                    throw new UserError("Comparison cancelled.");
                }
                // Discard helper output, avoiding confidential text in diagnostics and pipe deadlock.
                process.OutputDataReceived += delegate { };
                process.ErrorDataReceived += delegate { };
                process.BeginOutputReadLine();
                process.BeginErrorReadLine();
                process.WaitForExit();
                TryDelete(Path.Combine(workingDirectory, "ocr.pid"));
                TryDelete(Path.Combine(workingDirectory, "ocr.started"));
                TryDelete(Path.Combine(workingDirectory, "ocr.process"));
                if (process.ExitCode != 0 || !File.Exists(output))
                    throw new UserError("Local OCR failed for the " + label + " PDF. Check that it is unencrypted and that Windows has an OCR language installed (Settings > Time & language > Language & region > Language options). No document text was sent to a service.");
            }
            if (!File.Exists(status)) throw new UserError("Local OCR did not produce a complete status record. Please try again.");
            OcrStatus ocrStatus = JsonFile.Read<OcrStatus>(status);
            if (ocrStatus == null || ocrStatus.State != "Complete" || ocrStatus.PagesCompleted != ocrStatus.PageCount)
                throw new UserError("Local OCR did not confirm that every PDF page was processed. Please try again.");
            if (ocrStatus.BlankPages != null && ocrStatus.BlankPages.Length > 0)
                throw new UserError("Local OCR found pages without recognised words in the " + label + " PDF. Review those pages and use a checked local Word transcription, or remove genuinely blank pages from a copy before retrying. A crossover figure would risk omitting scanned content.");
            return output;
        }

        private static string QuoteArgument(string value)
        {
            // Windows CommandLineToArgvW quoting; arguments are passed directly, never through a shell command.
            var builder = new StringBuilder("\"");
            int slashes = 0;
            foreach (char c in value)
            {
                if (c == '\\') { slashes++; continue; }
                if (c == '"') { builder.Append('\\', slashes * 2 + 1); builder.Append(c); slashes = 0; continue; }
                builder.Append('\\', slashes); slashes = 0; builder.Append(c);
            }
            builder.Append('\\', slashes * 2); builder.Append('"');
            return builder.ToString();
        }

        private static void WriteProcessIdentity(string workingDirectory, string prefix, int pid)
        {
            long startTicks;
            using (Process process = Process.GetProcessById(pid))
                startTicks = process.StartTime.ToUniversalTime().Ticks;
            string temporaryRecord = Path.Combine(workingDirectory, prefix + ".process.tmp");
            JsonFile.Write(temporaryRecord, new { Id = pid, StartTimeUtcTicks = startTicks });
            File.Move(temporaryRecord, Path.Combine(workingDirectory, prefix + ".process"));
            File.WriteAllText(Path.Combine(workingDirectory, prefix + ".started"), startTicks.ToString(CultureInfo.InvariantCulture), Encoding.ASCII);
            // Write PID last, so a reader never sees an identity before its start-time check exists.
            File.WriteAllText(Path.Combine(workingDirectory, prefix + ".pid"), pid.ToString(CultureInfo.InvariantCulture), Encoding.ASCII);
        }

        private static void SaveDocx(dynamic document, string path)
        {
            if (File.Exists(path)) throw new UserError("A working file already exists. Start a fresh comparison.");
            document.SaveAs2(FileName: path, FileFormat: 12, AddToRecentFiles: false, EmbedTrueTypeFonts: false,
                SaveNativePictureFormat: true, SaveFormsData: false, SaveAsAOCELetter: false);
        }

        private static string BuildSummary(ComparisonRequest request, ComparisonResult result)
        {
            var text = new StringBuilder();
            if (result.Sections != null)
            {
                foreach (ComparisonSectionResult section in result.Sections)
                {
                    if (text.Length > 0) text.AppendLine();
                    text.AppendLine(section.Scope);
                    AppendFigures(text, section.CrossoverPercent, section.NewWords, section.InsertedWords);
                }
            }
            else AppendFigures(text, result.CrossoverPercent, result.NewWords, result.InsertedWords);
            if (result.Warnings != null && result.Warnings.Length > 0)
            {
                text.AppendLine();
                text.AppendLine("Review notes:");
                foreach (string warning in result.Warnings) text.AppendLine("- " + warning);
            }
            return text.ToString();
        }

        private static void AppendFigures(StringBuilder text, double? crossover, int newWords, int insertedWords)
        {
            text.AppendLine("Crossover: " + (crossover.HasValue ? crossover.Value.ToString("0.0", CultureInfo.InvariantCulture) + "%" : "Unavailable"));
            text.AppendLine("Total word count of new file: " + newWords.ToString(CultureInfo.InvariantCulture));
            text.AppendLine("Highlighted insertion word count: " + insertedWords.ToString(CultureInfo.InvariantCulture));
        }

        private static TextMetrics AnalyzeEpScope(string path, string workDirectory, string prefix, string scope)
        {
            string scopedPath = Path.Combine(workDirectory, prefix + "-scope-" + scope + ".docx");
            EpSections.ExtractScope(path, scopedPath, scope);
            return DocxContent.Analyze(scopedPath);
        }

        private static ComparisonSectionResult MeasureScope(string scope, TextMetrics source, TextMetrics compared, List<string> warnings)
        {
            var result = new ComparisonSectionResult { Scope = scope, NewWords = source.Words, InsertedWords = compared.InsertedWords };
            if (source.Words > 0 && compared.Words == source.Words && compared.TextFingerprint == source.TextFingerprint)
                result.CrossoverPercent = 100.0 * (source.Words - compared.InsertedWords) / source.Words;
            else if (source.Words > 0)
                warnings.Add(scope + " crossover was withheld because the final text represented by Word's comparison differs from the new text. Review the comparison manually.");
            return result;
        }

        private static bool IsOoxml(string path) { string extension = Path.GetExtension(path); return String.Equals(extension, ".docx", StringComparison.OrdinalIgnoreCase) || String.Equals(extension, ".docm", StringComparison.OrdinalIgnoreCase); }
        private static bool IsPdf(string path) { return String.Equals(Path.GetExtension(path), ".pdf", StringComparison.OrdinalIgnoreCase); }
        private static void Notify(Action<string> progress, string message) { try { if (progress != null) progress(message); } catch { } }
        private static void CheckCancelled(string workingDirectory) { if (File.Exists(Path.Combine(workingDirectory, "cancel.flag"))) throw new UserError("Comparison cancelled."); }
        private static void TryDelete(string path) { try { if (path != null) File.Delete(path); } catch { } }
        private static void Release(object value) { if (value != null && Marshal.IsComObject(value)) try { Marshal.FinalReleaseComObject(value); } catch { } }
        private static void CloseDocument(ref dynamic document)
        {
            if (document == null) return;
            try { document.Close(SaveChanges: 0); } catch { }
            Release((object)document);
            document = null;
        }
        private sealed class UserError : Exception { public UserError(string message) : base(message) { } }
        private sealed class OcrStatus
        {
            public OcrStatus() { }
            public string State { get; set; }
            public int PageCount { get; set; }
            public int PagesCompleted { get; set; }
            public int[] BlankPages { get; set; }
        }
    }
}
