using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Windows.Forms;

namespace LocalDocumentCompare
{
    internal sealed class BatchComparisonControl : UserControl
    {
        private readonly bool multiArchive;
        private readonly TextBox newDocument = new TextBox();
        private readonly CheckBox newDocumentOcr = new CheckBox();
        private readonly Button addArchives = new Button();
        private readonly TextBox folderPath = new TextBox();
        private readonly Button loadFiles = new Button();
        private readonly ComboBox jobType = new ComboBox();
        private readonly Label jobHint = new Label();
        private readonly Button resetBatch = new Button();
        private readonly DataGridView fileGrid = new DataGridView();
        private readonly Button moveUp = new Button();
        private readonly Button moveDown = new Button();
        private readonly Button removeFile = new Button();
        private readonly Label pairCount = new Label();
        private readonly TextBox outputPath = new TextBox();
        private readonly Button startBatch = new Button();
        private readonly Button cancelBatch = new Button();
        private readonly Button languageChecked = new Button();
        private readonly Button skipPair = new Button();
        private readonly Button retryReport = new Button();
        private readonly DataGridView resultsGrid = new DataGridView();
        private readonly TextBox details = new TextBox();
        private readonly Label state = new Label();
        private readonly Button openFolder = new Button();
        private readonly Button openReport = new Button();
        private readonly Button openComparison = new Button();
        private readonly Button saveResults = new Button();
        private readonly Panel configuration = new Panel();
        private readonly Timer timer = new Timer();
        private readonly List<BatchInput> files = new List<BatchInput>();
        private BatchPlan plan;
        private List<BatchRecord> records = new List<BatchRecord>();
        private int currentIndex = -1;
        private Process worker;
        private string jobDirectory;
        private string reportPath;
        private DateTime started;
        private bool busy;
        private bool reviewPaused;
        private bool reportBlocked;
        private bool stopBlocked;
        private bool keepJobDirectory;
        private bool cancelRequested;
        private bool pendingStart;
        private bool pendingStartConfirmed;
        private bool changingGrid;
        private bool nextScheduled;
        private bool disposing;
        private bool suppressLanguageNotice;
        private int runGeneration;
        private const string ReadyInstructions = "Use Browse or paste a local or shared-drive folder path, then select Load files. Only Word and PDF files directly in the folder are loaded. Earlier rows are archives; later rows are new files. Comparison and PDF OCR run on this PC. Results are saved in a new batch folder, with one subfolder per comparison and a live CSV report.";
        private const string MultiArchiveInstructions = "Use Add archives, drop archive files into the list, or load an archive folder, then choose one New document below. Each archive is compared against the same new document. Successful matches are ranked by crossover; EP uses Full specification crossover and shows all three scopes. Unavailable or unfinished results are unranked. Processing and PDF OCR stay on this PC; each comparison and a ranked CSV are saved in a new results folder.";
        private string Instructions { get { return multiArchive ? MultiArchiveInstructions : ReadyInstructions; } }

        public event EventHandler BusyChanged;
        public bool IsBusy { get { return busy; } }

        public BatchComparisonControl() : this(false) { }

        public BatchComparisonControl(bool multiArchiveMode)
        {
            multiArchive = multiArchiveMode;
            int extra = multiArchive ? 96 : 0;
            Font = BrandAppearance.Font(10F, FontStyle.Regular);
            ForeColor = BrandAppearance.Text;
            BackColor = Color.White;
            AutoScaleMode = AutoScaleMode.Dpi;
            Size = new Size(740, 960 + extra);
            Panel viewport = new Panel { Dock = DockStyle.Fill, AutoScroll = true };
            Controls.Add(viewport);
            TableLayoutPanel layout = new TableLayoutPanel
            {
                Dock = DockStyle.Top, Height = 960 + extra, ColumnCount = 1, RowCount = 8,
                Padding = new Padding(24, 12, 24, 14)
            };
            viewport.Controls.Add(layout);
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            foreach (int height in new[] { 42, 412 + extra, 92, 40, 104, 210, 46 })
                layout.RowStyles.Add(new RowStyle(SizeType.Absolute, height));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            layout.Controls.Add(new Label
            {
                Text = multiArchive ? "Find the best archive for one new document. Matches are ranked by crossover; EP uses Full specification." : "Compare every earlier file with each later file. Use 01, 02, 03 filename prefixes, then check the order below.",
                Dock = DockStyle.Fill, ForeColor = BrandAppearance.Muted,
                Font = BrandAppearance.Font(9F, FontStyle.Regular)
            }, 0, 0);

            configuration.Dock = DockStyle.Fill;
            configuration.Size = new Size(652, 412 + extra);
            configuration.Controls.Add(LabelAt("Job type", 0, 0));
            jobType.DropDownStyle = ComboBoxStyle.DropDownList;
            jobType.Items.AddRange(new object[] { "Full Document", "EP - description, full specification and claims" });
            jobType.SetBounds(92, 0, 452, 29);
            jobType.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            jobType.SelectedIndexChanged += delegate { UpdateJobHint(); };
            jobType.SelectedIndex = 0;
            configuration.Controls.Add(jobType);
            SetButton(resetBatch, "Reset", 92, false);
            resetBatch.Name = "resetBatch";
            resetBatch.AccessibleName = multiArchive ? "Reset Multi Archive" : "Reset batch comparison";
            resetBatch.AccessibleDescription = "Clear selected files and displayed results to start another batch. Saved files are kept.";
            resetBatch.SetBounds(560, 0, 92, 31);
            resetBatch.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            resetBatch.Click += delegate { ResetBatch(); };
            configuration.Controls.Add(resetBatch);
            jobHint.SetBounds(0, 34, 652, 38);
            jobHint.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            jobHint.Font = BrandAppearance.Font(9F, FontStyle.Regular);
            configuration.Controls.Add(jobHint);
            configuration.Controls.Add(LabelAt(multiArchive ? "Folder of archives (optional)" : "Folder of documents", 0, 80));
            folderPath.SetBounds(0, 106, 447, 29);
            folderPath.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            folderPath.TextChanged += delegate
            {
                if (busy) return;
                files.Clear(); RefreshFiles(-1);
                state.Text = multiArchive ? "The folder changed. Load files or add archive documents before searching." : "The folder changed. Load files and check their order before starting.";
            };
            configuration.Controls.Add(folderPath);
            Button browseFolder = ButtonAt("Browse…", 454, 105, 92, 31, true);
            browseFolder.Name = "browseFolder";
            browseFolder.AccessibleName = "Browse for a folder of documents";
            browseFolder.AccessibleDescription = "Choose a local or shared-drive folder, then select Load files.";
            browseFolder.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            browseFolder.Click += delegate
            {
                using (FolderBrowserDialog picker = new FolderBrowserDialog { Description = multiArchive ? "Choose a local or shared-drive folder of archive documents" : "Choose a local or shared-drive folder of archive and new documents", SelectedPath = folderPath.Text })
                    if (picker.ShowDialog(FindForm()) == DialogResult.OK) folderPath.Text = picker.SelectedPath;
            };
            configuration.Controls.Add(browseFolder);
            loadFiles.Text = "Load files";
            loadFiles.SetBounds(553, 105, 99, 31);
            loadFiles.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            BrandAppearance.StyleButton(loadFiles, false);
            loadFiles.Click += LoadFiles;
            configuration.Controls.Add(loadFiles);
            ConfigureGrid(fileGrid, true);
            fileGrid.SetBounds(0, 145, 652, 152);
            fileGrid.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            fileGrid.Columns.Add(new DataGridViewTextBoxColumn { Name = "order", HeaderText = "Order", Width = 54, ReadOnly = true, SortMode = DataGridViewColumnSortMode.NotSortable });
            fileGrid.Columns.Add(new DataGridViewTextBoxColumn { Name = "file", HeaderText = multiArchive ? "Archive file" : "File", AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill, ReadOnly = true, SortMode = DataGridViewColumnSortMode.NotSortable });
            fileGrid.Columns.Add(new DataGridViewCheckBoxColumn { Name = "ocr", HeaderText = "PDF OCR", Width = 85, SortMode = DataGridViewColumnSortMode.NotSortable });
            fileGrid.CurrentCellDirtyStateChanged += delegate { if (fileGrid.IsCurrentCellDirty) fileGrid.CommitEdit(DataGridViewDataErrorContexts.Commit); };
            fileGrid.CellValueChanged += delegate(object sender, DataGridViewCellEventArgs e)
            {
                if (changingGrid || e.RowIndex < 0 || e.RowIndex >= files.Count || e.ColumnIndex != 2) return;
                files[e.RowIndex].UseOcr = IsPdf(files[e.RowIndex].Path) && Convert.ToBoolean(fileGrid.Rows[e.RowIndex].Cells[2].Value);
            };
            fileGrid.SelectionChanged += delegate { UpdateOrderingButtons(); };
            configuration.Controls.Add(fileGrid);
            SetButton(moveUp, "Move up", 95, false);
            SetButton(moveDown, "Move down", 105, false);
            SetButton(removeFile, "Remove", 90, false);
            FlowLayoutPanel ordering = new FlowLayoutPanel { Location = new Point(0, 301), Size = new Size(652, 37), Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right, WrapContents = false };
            ordering.Controls.AddRange(new Control[] { moveUp, moveDown, removeFile });
            if (multiArchive)
            {
                moveUp.Visible = moveDown.Visible = false;
                SetButton(addArchives, "Add archives…", 148, true);
                addArchives.Name = "addArchives";
                addArchives.Click += AddArchiveFiles;
                ordering.Controls.Add(addArchives);
                ordering.Controls.SetChildIndex(addArchives, 0);
                fileGrid.AllowDrop = true;
                fileGrid.DragEnter += delegate(object sender, DragEventArgs e)
                { if (!busy && e.Data.GetDataPresent(DataFormats.FileDrop)) e.Effect = DragDropEffects.Copy; };
                fileGrid.DragDrop += delegate(object sender, DragEventArgs e)
                { if (!busy && e.Data.GetDataPresent(DataFormats.FileDrop)) AddArchivePaths((string[])e.Data.GetData(DataFormats.FileDrop)); };
            }
            moveUp.Click += delegate { MoveSelected(-1); };
            moveDown.Click += delegate { MoveSelected(1); };
            removeFile.Click += delegate
            {
                int selected = SelectedFile();
                if (selected >= 0) { files.RemoveAt(selected); RefreshFiles(Math.Min(selected, files.Count - 1)); }
            };
            configuration.Controls.Add(ordering);
            pairCount.SetBounds(311, 310, 340, 25);
            pairCount.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            pairCount.TextAlign = ContentAlignment.MiddleRight;
            pairCount.ForeColor = BrandAppearance.Muted;
            configuration.Controls.Add(pairCount);
            configuration.Controls.Add(LabelAt(multiArchive ? "Save Multi Archive results to" : "Save batch to", 0, 344));
            outputPath.SetBounds(0, 372, 447, 29);
            outputPath.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            outputPath.Text = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "LocalDocumentCompare", "Results");
            configuration.Controls.Add(outputPath);
            Button desktop = ButtonAt("Desktop", 454, 371, 92, 31, false);
            desktop.Name = "saveDesktop";
            desktop.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            desktop.Click += delegate { outputPath.Text = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory); };
            configuration.Controls.Add(desktop);
            Button browseOutput = ButtonAt("Browse…", 553, 371, 99, 31, false);
            browseOutput.Name = "browseOutput";
            browseOutput.AccessibleName = "Browse for a batch results folder";
            browseOutput.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            browseOutput.Click += delegate
            {
                using (FolderBrowserDialog picker = new FolderBrowserDialog { Description = "Choose a Desktop, local or shared-drive folder for batch results", SelectedPath = outputPath.Text })
                    if (picker.ShowDialog(FindForm()) == DialogResult.OK) outputPath.Text = picker.SelectedPath;
            };
            configuration.Controls.Add(browseOutput);
            if (multiArchive)
            {
                foreach (Control control in configuration.Controls)
                    if (control.Top >= 344) control.Top += extra;
                configuration.Controls.Add(LabelAt("New document", 0, 344));
                newDocument.Name = "multiArchiveNew";
                newDocument.SetBounds(0, 370, 544, 29);
                newDocument.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
                newDocument.AllowDrop = true;
                newDocument.TextChanged += delegate
                {
                    if (busy) return;
                    newDocumentOcr.Checked = false;
                    newDocumentOcr.Enabled = IsPdf(newDocument.Text.Trim().Trim('"'));
                    UpdateJobHint(); UpdateActions();
                };
                newDocument.DragEnter += delegate(object sender, DragEventArgs e)
                {
                    if (!busy && e.Data.GetDataPresent(DataFormats.FileDrop) && ((string[])e.Data.GetData(DataFormats.FileDrop)).Length == 1)
                        e.Effect = DragDropEffects.Copy;
                };
                newDocument.DragDrop += delegate(object sender, DragEventArgs e)
                {
                    if (busy || !e.Data.GetDataPresent(DataFormats.FileDrop)) return;
                    string[] dropped = (string[])e.Data.GetData(DataFormats.FileDrop);
                    if (dropped.Length == 1) newDocument.Text = dropped[0];
                };
                configuration.Controls.Add(newDocument);
                Button browseNew = ButtonAt("Browse…", 553, 369, 99, 31, true);
                browseNew.Name = "browseMultiNew";
                browseNew.AccessibleName = "Browse for the new document";
                browseNew.Anchor = AnchorStyles.Top | AnchorStyles.Right;
                browseNew.Click += delegate
                {
                    using (OpenFileDialog picker = new OpenFileDialog { Title = "Choose the new document", Filter = "Word and PDF documents|*.docx;*.doc;*.pdf", CheckFileExists = true })
                        if (picker.ShowDialog(FindForm()) == DialogResult.OK) newDocument.Text = picker.FileName;
                };
                configuration.Controls.Add(browseNew);
                newDocumentOcr.Text = "New PDF has scanned pages — use local OCR";
                newDocumentOcr.Name = "multiArchiveNewOcr";
                newDocumentOcr.SetBounds(0, 406, 652, 25);
                newDocumentOcr.Enabled = false;
                newDocumentOcr.Font = BrandAppearance.Font(9F, FontStyle.Regular);
                configuration.Controls.Add(newDocumentOcr);
                jobType.TabIndex = 0;
                resetBatch.TabIndex = 1;
                folderPath.TabIndex = 2;
                browseFolder.TabIndex = 3;
                loadFiles.TabIndex = 4;
                fileGrid.TabIndex = 5;
                ordering.TabIndex = 6;
                addArchives.TabIndex = 0;
                removeFile.TabIndex = 1;
                newDocument.TabIndex = 7;
                browseNew.TabIndex = 8;
                newDocumentOcr.TabIndex = 9;
                outputPath.TabIndex = 10;
                desktop.TabIndex = 11;
                browseOutput.TabIndex = 12;
            }
            layout.Controls.Add(configuration, 0, 1);

            FlowLayoutPanel actions = new FlowLayoutPanel { Dock = DockStyle.Fill, Padding = new Padding(0, 5, 0, 0), WrapContents = true };
            SetButton(startBatch, multiArchive ? "Find best match" : "Start batch", 150, true);
            SetButton(cancelBatch, multiArchive ? "Cancel search" : "Cancel batch", 112, false);
            SetButton(languageChecked, "Language checked — continue", 256, false);
            SetButton(skipPair, multiArchive ? "Skip archive" : "Skip pair", multiArchive ? 108 : 92, false);
            SetButton(retryReport, "Retry saving results", 168, false);
            foreach (Button button in new[] { startBatch, cancelBatch, languageChecked, skipPair, retryReport }) button.Height = 40;
            actions.Controls.AddRange(new Control[] { startBatch, cancelBatch, languageChecked, skipPair, retryReport });
            startBatch.Click += StartBatch;
            cancelBatch.Click += delegate { CancelAndClose(); };
            languageChecked.Click += delegate { if (busy && reviewPaused && !reportBlocked && !stopBlocked) StartCurrentPair(true); };
            skipPair.Click += delegate { SkipCurrentPair(); };
            retryReport.Click += delegate { RetryReport(); };
            layout.Controls.Add(actions, 0, 2);
            state.Text = multiArchive ? "Add archives, choose the new document and select Full Document or EP." : "Browse to a folder, load files, check the order and choose Full Document or EP.";
            state.Dock = DockStyle.Fill;
            state.TextAlign = ContentAlignment.MiddleLeft;
            layout.Controls.Add(state, 0, 3);
            details.Multiline = true;
            details.ReadOnly = true;
            details.BorderStyle = BorderStyle.None;
            details.BackColor = BrandAppearance.Surface;
            details.ForeColor = BrandAppearance.Muted;
            details.Font = BrandAppearance.Font(9F, FontStyle.Regular);
            details.ScrollBars = ScrollBars.Vertical;
            details.Dock = DockStyle.Fill;
            details.Text = Instructions;
            layout.Controls.Add(details, 0, 4);
            ConfigureGrid(resultsGrid, false);
            resultsGrid.Dock = DockStyle.Fill;
            resultsGrid.ReadOnly = true;
            resultsGrid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            AddResultColumn("new", "New file", 20F, 100);
            AddResultColumn("archive", "Archive file", 20F, 100);
            AddResultColumn("scope", "Scope", 16F, 105);
            AddResultColumn("crossover", "Crossover", 11F, 75);
            AddResultColumn("newWords", "New words", 10F, 65);
            AddResultColumn("insertions", "Inserted words", 10F, 65);
            AddResultColumn("status", "Status", 13F, 95);
            if (multiArchive) resultsGrid.Columns.Insert(0, new DataGridViewTextBoxColumn
            { Name = "rank", HeaderText = "Rank", FillWeight = 6F, MinimumWidth = 45, SortMode = DataGridViewColumnSortMode.NotSortable });
            resultsGrid.SelectionChanged += delegate { UpdateOutputButtons(); };
            resultsGrid.CellDoubleClick += delegate { OpenSelectedComparison(); };
            layout.Controls.Add(resultsGrid, 0, 5);
            FlowLayoutPanel outputActions = new FlowLayoutPanel { Dock = DockStyle.Fill, Padding = new Padding(0, 5, 0, 0), WrapContents = true };
            SetButton(openComparison, "Open comparison", 154, false);
            SetButton(openFolder, multiArchive ? "Open results folder" : "Open batch folder", 155, false);
            SetButton(openReport, "Open CSV report", 146, false);
            SetButton(saveResults, "Save results table", 160, false);
            outputActions.Controls.AddRange(new Control[] { openComparison, openFolder, openReport, saveResults });
            openComparison.Click += delegate { OpenSelectedComparison(); };
            openFolder.Click += delegate { if (plan != null) OpenPath(plan.OutputDirectory); };
            openReport.Click += delegate { OpenPath(reportPath); };
            saveResults.Click += delegate
            {
                if (busy) { RetryReport(); return; }
                if (plan == null) return;
                if (SaveReport()) state.Text = "The results table was saved.";
                else state.Text = "The results table could not be saved. Close the CSV or check the destination, then retry.";
            };
            layout.Controls.Add(outputActions, 0, 6);
            timer.Interval = 350;
            timer.Tick += PollWorker;
            configuration.Resize += delegate { LayoutConfiguration(); };
            LayoutConfiguration();
            RefreshFiles(-1);
            UpdateActions();
        }

        private void LayoutConfiguration()
        {
            int width = configuration.ClientSize.Width;
            if (width < 220) return;
            jobType.Width = Math.Max(1, width - jobType.Left - 108);
            resetBatch.Left = width - resetBatch.Width;
            jobHint.Width = fileGrid.Width = width;
            folderPath.Width = outputPath.Width = width - 205;
            if (multiArchive)
            {
                newDocument.Width = width - 108;
                newDocumentOcr.Width = width;
                Control[] browseNew = configuration.Controls.Find("browseMultiNew", false);
                if (browseNew.Length == 1) browseNew[0].Left = width - browseNew[0].Width;
            }
            loadFiles.Left = width - loadFiles.Width;
            foreach (string name in new[] { "browseFolder", "saveDesktop" })
            {
                Control[] found = configuration.Controls.Find(name, false);
                if (found.Length == 1) found[0].Left = width - 198;
            }
            Control[] output = configuration.Controls.Find("browseOutput", false);
            if (output.Length == 1) output[0].Left = width - output[0].Width;
            pairCount.Left = Math.Max(311, width - pairCount.Width);
            foreach (Control control in configuration.Controls)
                if (control is FlowLayoutPanel) control.Width = width;
        }

        private static Label LabelAt(string text, int x, int y)
        {
            return new Label { Text = text, AutoSize = true, Location = new Point(x, y), Font = BrandAppearance.Font(10F, FontStyle.Bold) };
        }

        private void ResetBatch()
        {
            if (busy || worker != null || disposing || keepJobDirectory) return;
            timer.Stop();
            runGeneration++;
            nextScheduled = false;
            pendingStart = pendingStartConfirmed = false;
            reviewPaused = reportBlocked = stopBlocked = cancelRequested = false;
            suppressLanguageNotice = false;
            started = DateTime.MinValue;
            currentIndex = -1;
            plan = null;
            reportPath = null;
            records = new List<BatchRecord>();
            files.Clear();
            folderPath.Clear();
            newDocument.Clear(); newDocumentOcr.Checked = false; newDocumentOcr.Enabled = false;
            RefreshFiles(-1);
            RefreshResults();
            UpdateActions();
            state.Text = multiArchive ? "Ready — add archives, then choose the new document below." : "Ready — browse to a folder and load the files.";
            details.Text = Instructions;
            if (!String.IsNullOrEmpty(jobDirectory))
                details.AppendText("\r\n\r\nPrevious temporary files remain: " + jobDirectory);
            if (multiArchive)
            {
                folderPath.Focus();
                Panel viewport = configuration.Parent.Parent as Panel;
                if (viewport != null) viewport.AutoScrollPosition = Point.Empty;
            }
            else folderPath.Focus();
        }

        private static void SetButton(Button button, string text, int width, bool primary)
        {
            button.Text = text; button.Size = new Size(width, 31);
            BrandAppearance.StyleButton(button, primary);
        }

        private static Button ButtonAt(string text, int x, int y, int width, int height, bool primary)
        {
            Button button = new Button(); SetButton(button, text, width, primary);
            button.SetBounds(x, y, width, height); return button;
        }

        private static void ConfigureGrid(DataGridView grid, bool editable)
        {
            grid.AllowUserToAddRows = false; grid.AllowUserToDeleteRows = false;
            grid.AllowUserToResizeRows = false; grid.AllowUserToOrderColumns = false;
            grid.RowHeadersVisible = false; grid.MultiSelect = false;
            grid.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            grid.BorderStyle = BorderStyle.None; grid.BackgroundColor = Color.White;
            grid.GridColor = BrandAppearance.Border;
            grid.CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal;
            grid.EnableHeadersVisualStyles = false;
            grid.ColumnHeadersHeight = 42;
            grid.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
            grid.ColumnHeadersDefaultCellStyle.BackColor = BrandAppearance.Purple;
            grid.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            grid.ColumnHeadersDefaultCellStyle.Font = BrandAppearance.Font(9F, FontStyle.Bold);
            grid.ColumnHeadersDefaultCellStyle.WrapMode = DataGridViewTriState.True;
            grid.DefaultCellStyle.Font = BrandAppearance.Font(9F, FontStyle.Regular);
            grid.DefaultCellStyle.ForeColor = BrandAppearance.Text;
            grid.DefaultCellStyle.SelectionBackColor = Color.FromArgb(237, 228, 247);
            grid.DefaultCellStyle.SelectionForeColor = BrandAppearance.Text;
            grid.RowTemplate.Height = 28;
            grid.ReadOnly = !editable;
        }

        private void AddResultColumn(string name, string title, float weight, int minimum)
        {
            resultsGrid.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = name, HeaderText = title, FillWeight = weight, MinimumWidth = minimum,
                SortMode = DataGridViewColumnSortMode.NotSortable
            });
        }

        private void LoadFiles(object sender, EventArgs e)
        {
            if (busy) return;
            try
            {
                List<BatchInput> loaded = BatchPlan.LoadFolder(folderPath.Text.Trim().Trim('"'));
                if (multiArchive && !String.IsNullOrWhiteSpace(newDocument.Text))
                {
                    string newPath = Path.GetFullPath(newDocument.Text.Trim().Trim('"'));
                    loaded.RemoveAll(input => String.Equals(input.Path, newPath, StringComparison.OrdinalIgnoreCase));
                }
                files.Clear(); files.AddRange(loaded); RefreshFiles(files.Count > 0 ? 0 : -1);
                state.Text = multiArchive
                    ? files.Count >= 1 ? "Check the archives and PDF OCR choices, then find the best match. The selected new file is excluded from folder loading." : "The folder has no archive documents to compare. Choose another folder or add archives."
                    : files.Count >= 2 ? "Check the file order and PDF OCR choices, then start the batch." : "The folder needs at least two Word or PDF documents.";
            }
            catch (Exception error)
            {
                state.Text = "The folder could not be loaded.";
                details.Text = error.Message;
            }
        }

        private void AddArchiveFiles(object sender, EventArgs e)
        {
            if (busy || !multiArchive) return;
            using (OpenFileDialog picker = new OpenFileDialog
            { Title = "Choose archive documents", Filter = "Word and PDF documents|*.docx;*.doc;*.pdf", Multiselect = true, CheckFileExists = true })
                if (picker.ShowDialog(FindForm()) == DialogResult.OK) AddArchivePaths(picker.FileNames);
        }

        private void AddArchivePaths(IEnumerable<string> paths)
        {
            if (busy || !multiArchive) return;
            try
            {
                var added = new List<BatchInput>();
                var seen = new HashSet<string>(files.Select(input => input.Path), StringComparer.OrdinalIgnoreCase);
                foreach (string selected in paths)
                {
                    ComparisonEngine.RequireSourcePath(selected);
                    string path = Path.GetFullPath(selected);
                    string extension = Path.GetExtension(path);
                    if (!(String.Equals(extension, ".docx", StringComparison.OrdinalIgnoreCase) ||
                        String.Equals(extension, ".doc", StringComparison.OrdinalIgnoreCase) || IsPdf(path)) ||
                        Path.GetFileName(path).StartsWith("~$", StringComparison.Ordinal) || !File.Exists(path))
                        throw new InvalidOperationException("Choose readable DOCX, DOC or PDF archive files. Word lock files cannot be compared.");
                    if (seen.Add(path)) added.Add(new BatchInput { Path = path });
                }
                files.AddRange(added); RefreshFiles(files.Count > 0 ? 0 : -1);
                state.Text = added.Count + " archives added. Each archive will be compared with the same new document.";
            }
            catch (Exception error) { state.Text = "The archive selection could not be added."; details.Text = error.Message; }
        }

        private static bool IsPdf(string path)
        {
            try { return String.Equals(Path.GetExtension(path), ".pdf", StringComparison.OrdinalIgnoreCase); }
            catch (ArgumentException) { return false; }
        }

        private void UpdateJobHint()
        {
            bool ep = jobType.SelectedIndex == 1;
            bool laterPdf = multiArchive && IsPdf(newDocument.Text.Trim().Trim('"'));
            for (int index = 1; !multiArchive && index < files.Count; index++)
                if (IsPdf(files[index].Path)) { laterPdf = true; break; }
            jobHint.Text = ep
                ? (laterPdf ? (multiArchive ? "The new file is a PDF: choose Full Document for a reliable whole-document comparison." : "New files include PDFs: choose Full Document for a reliable whole-document comparison.") : "For EP PDFs or Word files converted from PDF, choose Full Document.") + "\r\nEP needs clear description and claims sections; converted sections can be unreliable."
                : "Full Document compares all text.\r\nRecommended for EP PDFs and Word files converted from PDF.";
            jobHint.ForeColor = ep && laterPdf ? BrandAppearance.Pink : BrandAppearance.Muted;
        }

        private int SelectedFile() { return fileGrid.SelectedRows.Count == 0 ? -1 : fileGrid.SelectedRows[0].Index; }

        private void MoveSelected(int offset)
        {
            if (busy) return;
            int selected = SelectedFile(), target = selected + offset;
            if (selected < 0 || target < 0 || target >= files.Count) return;
            BatchInput input = files[selected]; files[selected] = files[target]; files[target] = input;
            RefreshFiles(target);
        }

        private void RefreshFiles(int selected)
        {
            changingGrid = true;
            try
            {
                fileGrid.Rows.Clear();
                for (int index = 0; index < files.Count; index++)
                {
                    int row = fileGrid.Rows.Add(index + 1, Path.GetFileName(files[index].Path), files[index].UseOcr);
                    fileGrid.Rows[row].Cells[1].ToolTipText = files[index].Path;
                    if (!IsPdf(files[index].Path))
                    {
                        fileGrid.Rows[row].Cells[2].ReadOnly = true;
                        fileGrid.Rows[row].Cells[2].Style.BackColor = BrandAppearance.Surface;
                        fileGrid.Rows[row].Cells[2].ToolTipText = "OCR is available only for PDF inputs.";
                    }
                }
                fileGrid.ClearSelection();
                if (selected >= 0 && selected < fileGrid.Rows.Count)
                { fileGrid.Rows[selected].Selected = true; fileGrid.CurrentCell = fileGrid.Rows[selected].Cells[1]; }
            }
            finally { changingGrid = false; }
            long comparisons = multiArchive ? files.Count : (long)files.Count * Math.Max(0, files.Count - 1) / 2;
            pairCount.Text = files.Count + (multiArchive ? " archives · " : " files · ") + comparisons.ToString("N0") + " comparisons";
            UpdateJobHint();
            UpdateOrderingButtons(); UpdateActions();
        }

        private void UpdateOrderingButtons()
        {
            int selected = SelectedFile();
            moveUp.Enabled = !busy && selected > 0;
            moveDown.Enabled = !busy && selected >= 0 && selected < files.Count - 1;
            removeFile.Enabled = !busy && selected >= 0;
            addArchives.Enabled = !busy;
        }

        private void StartBatch(object sender, EventArgs e)
        {
            if (busy) return;
            try
            {
                fileGrid.EndEdit();
                if (files.Count < (multiArchive ? 1 : 2)) throw new InvalidOperationException(multiArchive ? "Choose the new document and add at least one archive." : "Load at least two files and check their order before starting.");
                if (Type.GetTypeFromProgID("Word.Application") == null) throw new InvalidOperationException("Install and activate desktop Microsoft Word before using this tool.");
                string type = jobType.SelectedIndex == 1 ? "EP" : "WF";
                plan = multiArchive
                    ? BatchPlan.CreateMultiArchive(files, new BatchInput { Path = newDocument.Text.Trim().Trim('"'), UseOcr = newDocumentOcr.Checked }, outputPath.Text.Trim().Trim('"'), type)
                    : BatchPlan.Create(files, outputPath.Text.Trim().Trim('"'), type);
                runGeneration++;
                records = new List<BatchRecord>();
                foreach (BatchPair pair in plan.Pairs) records.Add(new BatchRecord { Pair = pair, State = "Pending" });
                currentIndex = 0; reportPath = null; cancelRequested = false;
                suppressLanguageNotice = false;
                reviewPaused = false; reportBlocked = false; stopBlocked = false;
                pendingStart = true; pendingStartConfirmed = false;
                SetBusy(true); RefreshResults();
                details.Text = (multiArchive ? "Multi Archive results: " : "Batch results: ") + plan.OutputDirectory + "\r\nEach comparison runs separately and keeps its completed Word document and summary.";
                if (SaveReport()) StartCurrentPair(false);
            }
            catch (Exception error)
            {
                state.Text = multiArchive ? "The archive search could not start." : "The batch could not start.";
                details.Text = error.Message;
                SetBusy(false);
            }
        }

        private void SetBusy(bool value)
        {
            bool changed = busy != value;
            busy = value; configuration.Enabled = !value;
            UpdateActions(); UpdateOrderingButtons();
            if (changed && BusyChanged != null) BusyChanged(this, EventArgs.Empty);
        }

        private void UpdateActions()
        {
            resetBatch.Enabled = !busy && worker == null && !keepJobDirectory && !disposing;
            startBatch.Enabled = !busy && files.Count >= (multiArchive ? 1 : 2) && (!multiArchive || !String.IsNullOrWhiteSpace(newDocument.Text));
            cancelBatch.Enabled = busy;
            cancelBatch.Text = stopBlocked ? "Retry stop" : multiArchive ? "Cancel search" : "Cancel batch";
            languageChecked.Visible = busy && reviewPaused && !reportBlocked && !stopBlocked;
            skipPair.Visible = languageChecked.Visible;
            retryReport.Visible = busy && reportBlocked && !stopBlocked;
            UpdateOutputButtons();
        }

        private void UpdateOutputButtons()
        {
            // Keep selection/paint events independent of shared-drive latency.
            // Opening the selected path reports an unavailable destination itself.
            openFolder.Enabled = plan != null;
            openReport.Enabled = !busy && reportPath != null;
            saveResults.Enabled = plan != null && (!busy || reportBlocked) && worker == null && !stopBlocked;
            BatchRecord selected = SelectedResult();
            openComparison.Enabled = !busy && selected != null && selected.State == "Completed" && selected.Result != null && selected.Result.Success;
        }

        private void RefreshResults()
        {
            int selectedPair = SelectedResult() == null ? -1 : SelectedResult().Pair.Number;
            resultsGrid.Rows.Clear();
            string type = plan == null ? (jobType.SelectedIndex == 1 ? "EP" : "WF") : plan.JobType;
            Dictionary<int, int> ranks = multiArchive ? MultiArchiveRanking.Ranks(records, type) : null;
            if (multiArchive) resultsGrid.Columns["rank"].HeaderText = busy ? "Rank so far" : "Rank";
            foreach (BatchRecord record in multiArchive ? MultiArchiveRanking.Ordered(records, type) : records)
            {
                int value;
                int? rank = ranks != null && ranks.TryGetValue(record.Pair.Number, out value) ? (int?)value : null;
                bool completed = record.State == "Completed" && record.Result != null && record.Result.Success;
                if (completed && record.Result.Sections != null && record.Result.Sections.Length > 0)
                {
                    foreach (ComparisonSectionResult section in record.Result.Sections)
                        AddResultRow(record, section.Scope, section.CrossoverPercent, section.NewWords, section.InsertedWords, rank);
                }
                else AddResultRow(record, plan != null && plan.JobType == "EP" ? "EP" : "Whole document", completed ? record.Result.CrossoverPercent : null, completed ? (int?)record.Result.NewWords : null, completed ? (int?)record.Result.InsertedWords : null, rank);
            }
            resultsGrid.ClearSelection();
            foreach (DataGridViewRow row in resultsGrid.Rows)
            {
                BatchRecord record = row.Tag as BatchRecord;
                if (record != null && record.Pair.Number == selectedPair) { row.Selected = true; break; }
            }
            UpdateOutputButtons();
        }

        private void AddResultRow(BatchRecord record, string scope, double? crossover, int? words, int? insertions, int? rank)
        {
            var cells = new List<object>(new object[] { Path.GetFileName(record.Pair.New.Path), Path.GetFileName(record.Pair.Archive.Path), scope,
                crossover.HasValue ? crossover.Value.ToString("0.0") + "%" : record.State == "Completed" ? "Unavailable" : "",
                words.HasValue ? words.Value.ToString("N0") : "", insertions.HasValue ? insertions.Value.ToString("N0") : "", record.State });
            if (multiArchive) cells.Insert(0, rank.HasValue ? rank.Value.ToString() : "");
            int index = resultsGrid.Rows.Add(cells.ToArray());
            DataGridViewRow row = resultsGrid.Rows[index]; row.Tag = record;
            row.Cells["new"].ToolTipText = record.Pair.New.Path;
            row.Cells["archive"].ToolTipText = record.Pair.Archive.Path;
            if (record.Result != null && !String.IsNullOrEmpty(record.Result.Error)) row.Cells["status"].ToolTipText = record.Result.Error;
        }

        private BatchRecord SelectedResult()
        {
            return resultsGrid.SelectedRows.Count == 0 ? null : resultsGrid.SelectedRows[0].Tag as BatchRecord;
        }

        private bool SaveReport()
        {
            try
            {
                reportPath = BatchReport.Save(plan, records);
                reportBlocked = false; UpdateActions(); return true;
            }
            catch (Exception error)
            {
                reportBlocked = true;
                state.Text = busy ? (multiArchive ? "Archive search paused — the CSV report could not be saved." : "Batch paused — the CSV report could not be saved.") : "The CSV report could not be saved.";
                details.Text = "Completed comparison files are retained in " + plan.OutputDirectory + ".\r\nClose the CSV in Excel, check that the destination is writable, then choose " + (busy ? "Retry saving results" : "Save results table") + ".\r\n" + error.Message;
                UpdateActions(); return false;
            }
        }

        private void RetryReport()
        {
            if (!busy || !reportBlocked || stopBlocked || !SaveReport()) return;
            if (reviewPaused) ShowReview();
            else if (pendingStart) StartCurrentPair(pendingStartConfirmed);
            else ScheduleNext();
        }

        private void StartCurrentPair(bool confirmed)
        {
            if (!busy || worker != null || currentIndex < 0 || currentIndex >= records.Count || stopBlocked) return;
            BatchRecord record = records[currentIndex];
            reviewPaused = false; pendingStart = true; pendingStartConfirmed = confirmed;
            record.State = "Running"; record.Result = null; RefreshResults(); UpdateActions();
            if (!SaveReport()) { record.State = "Pending"; RefreshResults(); return; }
            try
            {
                jobDirectory = Path.Combine(Path.GetTempPath(), "LocalDocumentCompare", Guid.NewGuid().ToString("N"));
                ComparisonEngine.RequireLocalPath(jobDirectory);
                Directory.CreateDirectory(jobDirectory); keepJobDirectory = false;
                ComparisonRequest request = new ComparisonRequest
                {
                    ArchivePath = record.Pair.Archive.Path, NewPath = record.Pair.New.Path,
                    OutputPath = record.Pair.OutputPath, WorkDirectory = jobDirectory,
                    OcrArchive = record.Pair.Archive.UseOcr, OcrNew = record.Pair.New.UseOcr,
                    JobType = plan.JobType, MatchingLanguageConfirmed = confirmed
                };
                string requestPath = Path.Combine(jobDirectory, "request.json");
                JsonFile.Write(requestPath, request);
                ProcessStartInfo start = new ProcessStartInfo(Assembly.GetExecutingAssembly().Location,
                    "--worker " + Program.Quote(requestPath) + " " + Program.Quote(Path.Combine(jobDirectory, "result.json")))
                {
                    UseShellExecute = false, CreateNoWindow = true, WindowStyle = ProcessWindowStyle.Hidden,
                    WorkingDirectory = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location)
                };
                worker = Process.Start(start); started = DateTime.UtcNow;
                if (worker == null) throw new InvalidOperationException("The comparison worker could not be started.");
                pendingStart = false;
                state.Text = PairLabel(record) + " — starting comparison.";
                details.Text = "New: " + record.Pair.New.Path + "\r\nArchive: " + record.Pair.Archive.Path + "\r\nCompleted results stay in " + plan.OutputDirectory + ".";
                timer.Start();
            }
            catch (Exception error)
            {
                pendingStart = false;
                record.State = "Failed";
                record.Result = new ComparisonResult { Success = false, Error = "The pair could not start: " + error.Message };
                string leftover = CleanupJob();
                if (leftover != null) record.Result.Warnings = new[] { "Temporary files remain: " + leftover };
                RefreshResults();
                if (SaveReport()) ScheduleNext();
            }
        }

        private string PairLabel(BatchRecord record) { return (multiArchive ? "Archive " : "Pair ") + record.Pair.Number + " of " + records.Count; }

        private void PollWorker(object sender, EventArgs e)
        {
            if (cancelRequested) { CancelAndClose(); return; }
            if (worker == null || !busy) return;
            try
            {
                if (!worker.HasExited)
                {
                    if ((DateTime.UtcNow - started).TotalMinutes >= 20)
                    {
                        if (!StopWorker()) { HoldForStop(); return; }
                        FinishPair(new ComparisonResult { Success = false, Error = "The comparison exceeded 20 minutes and was stopped. Review this pair in Word before retrying." }, false);
                        return;
                    }
                    string path = Path.Combine(jobDirectory, "status.txt");
                    if (File.Exists(path))
                        try { state.Text = PairLabel(records[currentIndex]) + " — " + File.ReadAllText(path, Encoding.UTF8); }
                        catch (IOException) { }
                    return;
                }
                FinishPair(ReadResult(), false);
            }
            catch (Exception error)
            {
                if (!StopWorker()) { HoldForStop(); return; }
                FinishPair(new ComparisonResult { Success = false, Error = "The comparison process could not finish: " + error.Message }, false);
            }
        }

        private ComparisonResult ReadResult()
        {
            try
            {
                string path = Path.Combine(jobDirectory, "result.json");
                if (File.Exists(path)) return JsonFile.Read<ComparisonResult>(path);
            }
            catch { }
            return new ComparisonResult { Success = false, Error = "The comparison stopped without a completed result. Open these files in Word and check for a conversion or protection prompt." };
        }

        private void FinishPair(ComparisonResult result, bool cancelled)
        {
            timer.Stop();
            if (result.LanguageReviewRequired && result.Error != null)
                result.Error = result.Error.Replace("Language checked — compare", "Language checked — continue");
            bool stopped = StopOwnedChildren();
            if (worker != null) { worker.Dispose(); worker = null; }
            BatchRecord record = records[currentIndex]; record.Result = result;
            reviewPaused = !cancelled && !cancelRequested && !result.Success && result.LanguageReviewRequired;
            record.State = result.Success ? "Completed" : cancelled || cancelRequested ? "Cancelled" : reviewPaused ? "Review required" : "Failed";
            pendingStart = false;
            string leftover = CleanupJob();
            if (leftover != null)
            {
                List<string> notes = new List<string>(result.Warnings ?? new string[0]);
                notes.Add("Temporary files remain: " + leftover); result.Warnings = notes.ToArray();
            }
            RefreshResults();
            if (!stopped) { HoldForStop(); SaveReport(); return; }
            if (!SaveReport()) return;
            if (cancelled || cancelRequested) FinishCancellation();
            else if (reviewPaused) ShowReview();
            else ScheduleNext();
        }

        private void ShowReview()
        {
            BatchRecord record = records[currentIndex];
            state.Text = PairLabel(record) + " — language check required; " + (multiArchive ? "search paused." : "batch paused.");
            details.Text = "New: " + record.Pair.New.Path + "\r\nArchive: " + record.Pair.Archive.Path + "\r\n\r\n" + record.Result.Error + "\r\n\r\nCheck this pair, then choose Language checked — continue. This confirms only this pair. Choose " + (multiArchive ? "Skip archive" : "Skip pair") + " to record a failure and continue with the remaining comparisons.";
            UpdateActions();
        }

        private void SkipCurrentPair()
        {
            if (!busy || !reviewPaused || worker != null || reportBlocked || stopBlocked) return;
            BatchRecord record = records[currentIndex];
            record.State = "Failed";
            record.Result = new ComparisonResult { Success = false, Error = "Skipped because the source language was not confirmed manually." };
            reviewPaused = false; RefreshResults(); UpdateActions();
            if (SaveReport()) ScheduleNext();
        }

        private void ScheduleNext()
        {
            if (!busy || nextScheduled || reviewPaused || reportBlocked || stopBlocked || cancelRequested) return;
            nextScheduled = true;
            int generation = runGeneration;
            BeginInvoke((Action)delegate
            {
                if (generation != runGeneration) return;
                nextScheduled = false;
                if (disposing || !busy || reviewPaused || reportBlocked || stopBlocked || cancelRequested) return;
                currentIndex++;
                if (currentIndex >= records.Count) CompleteBatch(false);
                else StartCurrentPair(false);
            });
        }

        private void CompleteBatch(bool cancelled, bool notifyLanguageMismatch = true)
        {
            bool finishedRun = busy;
            timer.Stop(); reviewPaused = false; reportBlocked = false; stopBlocked = false;
            runGeneration++;
            pendingStart = false; nextScheduled = false;
            int completed = 0, failed = 0, notRun = 0;
            foreach (BatchRecord record in records)
            {
                if (record.State == "Completed") completed++;
                else if (record.State == "Failed" || record.State == "Cancelled") failed++;
                else notRun++;
            }
            state.Text = multiArchive ? (cancelled ? "Archive search cancelled — completed results retained." : "Archive search finished.") : cancelled ? "Batch cancelled — completed results retained." : "Batch finished.";
            details.Text = completed + " completed; " + failed + " failed or cancelled; " + notRun + " not run.\r\nResults: " + plan.OutputDirectory + "\r\nCSV report: " + (reportPath ?? "could not be saved") + "\r\nReview Word markup and any PDF/OCR or language notes before relying on the figures.";
            if (multiArchive)
            {
                List<BatchRecord> ranked = MultiArchiveRanking.Ordered(records, plan.JobType).Where(r => MultiArchiveRanking.Score(r, plan.JobType).HasValue).ToList();
                if (ranked.Count == 0) details.Text = "No archive has a verified crossover available to rank. Review the statuses and notes.\r\n\r\n" + details.Text;
                else
                {
                    double best = MultiArchiveRanking.Score(ranked[0], plan.JobType).Value;
                    List<BatchRecord> winners = ranked.Where(r => MultiArchiveRanking.Score(r, plan.JobType).Value == best).ToList();
                    string names = String.Join(", ", winners.Take(5).Select(r => Path.GetFileName(r.Pair.Archive.Path)));
                    if (winners.Count > 5) names += " and " + (winners.Count - 5) + " more";
                    details.Text = (winners.Count == 1 ? "Best match" : "Joint best matches") + (cancelled ? " among completed comparisons" : " among ranked archives") + ": " + names + " — " + best.ToString("0.0") + "% crossover" +
                        (plan.JobType == "EP" ? " (Full specification)." : ".") + "\r\n" + ranked.Count + " of " + records.Count + " archives ranked. Unavailable, failed or unfinished results have no rank.\r\n\r\n" + details.Text;
                }
            }
            SetBusy(false); RefreshResults();
            if (multiArchive)
            {
                Panel viewport = resultsGrid.Parent.Parent as Panel;
                if (viewport != null) viewport.ScrollControlIntoView(openComparison.Parent);
            }
            Form owner = FindForm();
            if (finishedRun && notifyLanguageMismatch && !suppressLanguageNotice && !disposing && !IsDisposed &&
                owner != null && owner.Visible && !owner.Disposing && !owner.IsDisposed)
            {
                if (multiArchive) LanguageMismatchNotice.ShowMultiArchive(owner, records);
                else LanguageMismatchNotice.ShowBatch(owner, records);
            }
        }

        public bool CancelAndClose()
        {
            return CancelAndClose(true);
        }

        internal bool CancelAndClose(bool notifyLanguageMismatch)
        {
            if (!busy) return true;
            if (!notifyLanguageMismatch) suppressLanguageNotice = true;
            cancelRequested = true; timer.Stop();
            if (!StopWorker()) { HoldForStop(); timer.Start(); return false; }
            if (worker != null)
            {
                ComparisonResult result = ReadResult();
                worker.Dispose(); worker = null;
                if (currentIndex >= 0 && currentIndex < records.Count)
                {
                    records[currentIndex].Result = result;
                    records[currentIndex].State = result.Success ? "Completed" : "Cancelled";
                }
            }
            else if (currentIndex >= 0 && currentIndex < records.Count)
            {
                BatchRecord current = records[currentIndex];
                if (current.State != "Completed" && current.State != "Failed")
                {
                    current.State = "Cancelled";
                    current.Result = new ComparisonResult { Success = false, Error = "The batch was cancelled before this pair completed." };
                }
            }
            CleanupJob();
            FinishCancellation(notifyLanguageMismatch); return true;
        }

        private void FinishCancellation(bool notifyLanguageMismatch = true)
        {
            foreach (BatchRecord record in records)
                if (record.State == "Pending" || record.State == "Running" || record.State == "Review required")
                { record.State = "Not run"; record.Result = null; }
            bool saved = SaveReport();
            CompleteBatch(true, notifyLanguageMismatch);
            if (!saved) details.AppendText("\r\nThe final CSV update failed. Completed files remain; close the CSV or check the destination before starting another batch. The previous report may not include the final cancellation status.");
        }

        private bool StopWorker()
        {
            try
            {
                if (worker != null && !worker.HasExited)
                {
                    File.WriteAllText(Path.Combine(jobDirectory, "cancel.flag"), "cancel");
                    if (!worker.WaitForExit(1200)) { worker.Kill(); if (!worker.WaitForExit(3000)) return false; }
                }
                if (worker != null && !worker.HasExited) return false;
            }
            catch { return false; }
            return StopOwnedChildren();
        }

        private bool StopOwnedChildren()
        {
            bool stopped = StopRecordedProcess("ocr", "powershell") & StopRecordedProcess("word", "WINWORD");
            keepJobDirectory = !stopped;
            return stopped;
        }

        private bool StopRecordedProcess(string name, string expectedName)
        {
            if (String.IsNullOrEmpty(jobDirectory)) return true;
            try
            {
                int pid; long ticks;
                string recordPath = Path.Combine(jobDirectory, name + ".process");
                if (File.Exists(recordPath))
                {
                    OwnedProcess record = JsonFile.Read<OwnedProcess>(recordPath);
                    pid = record.Id; ticks = record.StartTimeUtcTicks;
                }
                else
                {
                    string pidPath = Path.Combine(jobDirectory, name + ".pid");
                    string stampPath = Path.Combine(jobDirectory, name + ".started");
                    if (!File.Exists(pidPath) || !File.Exists(stampPath)) return true;
                    if (!Int32.TryParse(File.ReadAllText(pidPath).Trim(), out pid) || !Int64.TryParse(File.ReadAllText(stampPath).Trim(), out ticks)) return false;
                }
                using (Process owned = Process.GetProcessById(pid))
                {
                    if (!owned.HasExited && String.Equals(owned.ProcessName, expectedName, StringComparison.OrdinalIgnoreCase) && owned.StartTime.ToUniversalTime().Ticks == ticks)
                    { owned.Kill(); return owned.WaitForExit(3000); }
                }
                return true;
            }
            catch (ArgumentException) { return true; }
            catch { return false; }
        }

        private void HoldForStop()
        {
            stopBlocked = true;
            state.Text = multiArchive ? "The archive search is paused because its comparison session could not be stopped." : "The batch is paused because its comparison session could not be stopped.";
            details.Text = "Choose Retry stop before closing or starting another comparison. The app stops only its own recorded processes. Completed results remain in " + plan.OutputDirectory + ".\r\nTemporary files: " + jobDirectory;
            UpdateActions();
        }

        private string CleanupJob()
        {
            if (String.IsNullOrEmpty(jobDirectory)) return null;
            string path = jobDirectory;
            if (keepJobDirectory) return path;
            try
            {
                string root = Path.GetFullPath(Path.Combine(Path.GetTempPath(), "LocalDocumentCompare")) + Path.DirectorySeparatorChar;
                string resolved = Path.GetFullPath(path);
                if (!resolved.StartsWith(root, StringComparison.OrdinalIgnoreCase)) return path;
                if (Directory.Exists(resolved)) Directory.Delete(resolved, true);
                jobDirectory = null; return null;
            }
            catch { return path; }
        }

        private void OpenSelectedComparison()
        {
            BatchRecord selected = SelectedResult();
            if (!busy && selected != null && selected.State == "Completed" && selected.Result != null && selected.Result.Success)
                OpenPath(selected.Pair.OutputPath);
        }

        private void OpenPath(string path)
        {
            try { if (!String.IsNullOrEmpty(path)) Process.Start(new ProcessStartInfo(path) { UseShellExecute = true }); }
            catch { state.Text = "The saved result could not be opened. Browse to the batch folder to open it."; }
        }

        protected override void Dispose(bool disposingNow)
        {
            if (disposingNow)
            {
                disposing = true;
                if (busy) CancelAndClose(false);
                timer.Dispose();
            }
            base.Dispose(disposingNow);
        }

        private sealed class OwnedProcess
        {
            public OwnedProcess() { }
            public int Id { get; set; }
            public long StartTimeUtcTicks { get; set; }
        }
    }
}
