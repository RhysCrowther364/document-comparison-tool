using System;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;

namespace LocalDocumentCompare
{
    internal static class Program
    {
        [STAThread]
        private static int Main(string[] args)
        {
            if (args.Length == 3 && args[0] == "--worker")
            {
                ComparisonResult result;
                try
                {
                    ComparisonRequest request = JsonFile.Read<ComparisonRequest>(args[1]);
                    result = ComparisonEngine.Run(request, delegate(string message)
                    {
                        try { File.WriteAllText(Path.Combine(request.WorkDirectory, "status.txt"), message, Encoding.UTF8); }
                        catch (IOException) { } // UI polling must never interrupt a comparison.
                    });
                }
                catch (Exception)
                {
                    result = new ComparisonResult { Success = false, Error = "The comparison could not finish. Check that both files open in Word and that the destination is writable." };
                }
                JsonFile.Write(args[2], result);
                return result.Success ? 0 : 1;
            }
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.SetUnhandledExceptionMode(UnhandledExceptionMode.CatchException);
            Application.ThreadException += delegate
            {
                MessageBox.Show("An unexpected error occurred. Close and reopen Document Compare. Your source files are unchanged.", "Document Compare", MessageBoxButtons.OK, MessageBoxIcon.Error);
            };
            using (MainForm form = new MainForm())
            {
                if (args.Length == 2 && args[0] == "--ui-preview")
                {
                    form.Show();
                    Application.DoEvents();
                    using (Bitmap preview = new Bitmap(form.Width, form.Height))
                    {
                        form.DrawToBitmap(preview, new Rectangle(Point.Empty, form.Size));
                        preview.Save(args[1], System.Drawing.Imaging.ImageFormat.Png);
                    }
                    return 0;
                }
                Application.Run(form);
            }
            return 0;
        }

        internal static string Quote(string value)
        {
            // Windows CommandLineToArgvW quoting, including trailing backslashes.
            StringBuilder output = new StringBuilder("\"");
            int slashes = 0;
            foreach (char c in value)
            {
                if (c == '\\') { slashes++; continue; }
                if (c == '"') output.Append('\\', slashes * 2 + 1);
                else output.Append('\\', slashes);
                output.Append(c);
                slashes = 0;
            }
            output.Append('\\', slashes * 2).Append('"');
            return output.ToString();
        }
    }

    internal sealed class MainForm : Form
    {
        private readonly TextBox archive = new TextBox();
        private readonly TextBox revised = new TextBox();
        private readonly TextBox destination = new TextBox();
        private readonly CheckBox archiveOcr = new CheckBox();
        private readonly CheckBox newOcr = new CheckBox();
        private readonly ComboBox jobType = new ComboBox();
        private readonly Label jobHint = new Label();
        private readonly Button compare = new Button();
        private readonly Button cancel = new Button();
        private readonly Button reset = new Button();
        private readonly Button languageChecked = new Button();
        private readonly Button openDocument = new Button();
        private readonly Button openFolder = new Button();
        private readonly Label state = new Label();
        private readonly Label metrics = new Label();
        private readonly DataGridView sectionMetrics = new DataGridView();
        private readonly TextBox details = new TextBox();
        private readonly ProgressBar progress = new ProgressBar();
        private readonly Timer timer = new Timer();
        private readonly Panel inputs = new Panel();
        private readonly TabControl tabs = new TabControl();
        private readonly TabPage singleTab = new TabPage("Single comparison");
        private readonly TabPage batchTab = new TabPage("Batch comparison");
        private readonly BatchComparisonControl batch = new BatchComparisonControl();
        private readonly TabPage multiArchiveTab = new TabPage("Multi Archive");
        private readonly BatchComparisonControl multiArchive = new BatchComparisonControl(true);
        private Process worker;
        private string jobDirectory;
        private string lastOutput;
        private DateTime started;
        private bool cancelling;
        private bool keepJobDirectory;
        private readonly Image brandLogo;
        private readonly Icon brandIcon;

        public MainForm()
        {
            Text = "RWS | Document Compare";
            Font = BrandAppearance.Font(10F, FontStyle.Regular);
            BackColor = Color.White;
            ForeColor = BrandAppearance.Text;
            brandLogo = BrandAppearance.LoadLogo();
            brandIcon = BrandAppearance.LoadIcon();
            Icon = brandIcon;
            StartPosition = FormStartPosition.CenterScreen;
            ClientSize = new Size(820, Math.Min(864, Screen.PrimaryScreen.WorkingArea.Height - 90));
            MinimumSize = new Size(760, 580);
            AutoScaleMode = AutoScaleMode.Dpi;

            tabs.Dock = DockStyle.Fill;
            tabs.Padding = new Point(18, 8);
            singleTab.BackColor = Color.White;
            batchTab.BackColor = Color.White;
            multiArchiveTab.BackColor = Color.White;
            tabs.TabPages.Add(singleTab);
            tabs.TabPages.Add(batchTab);
            tabs.TabPages.Add(multiArchiveTab);
            Controls.Add(tabs);
            Panel viewport = new Panel { Dock = DockStyle.Fill, AutoScroll = true };
            singleTab.Controls.Add(viewport);
            TableLayoutPanel layout = new TableLayoutPanel { Dock = DockStyle.Top, Height = 884, Padding = new Padding(28, 22, 28, 18), ColumnCount = 1, RowCount = 8 };
            viewport.Controls.Add(layout);
            viewport.Resize += delegate { layout.Height = Math.Max(884, viewport.ClientSize.Height); };
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 100));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 360));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 58));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 30));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 8));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 140));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 58));

            Panel heading = CreateHeading();
            layout.Controls.Add(heading, 0, 0);

            batch.Dock = DockStyle.Fill;
            Panel batchHeading = new Panel { Dock = DockStyle.Top, Height = 122, Padding = new Padding(28, 22, 28, 0) };
            batchHeading.Controls.Add(CreateHeading());
            batchTab.Controls.Add(batch);
            batchTab.Controls.Add(batchHeading);
            batch.BusyChanged += delegate { UpdateModeAvailability(); };
            multiArchive.Dock = DockStyle.Fill;
            Panel multiHeading = new Panel { Dock = DockStyle.Top, Height = 122, Padding = new Padding(28, 22, 28, 0) };
            multiHeading.Controls.Add(CreateHeading());
            multiArchiveTab.Controls.Add(multiArchive);
            multiArchiveTab.Controls.Add(multiHeading);
            multiArchive.BusyChanged += delegate { UpdateModeAvailability(); };
            tabs.Selecting += delegate(object sender, TabControlCancelEventArgs e)
            {
                if ((worker != null && e.TabPage != singleTab) || (batch.IsBusy && e.TabPage != batchTab) ||
                    (multiArchive.IsBusy && e.TabPage != multiArchiveTab)) e.Cancel = true;
            };
            tabs.SelectedIndexChanged += delegate { AcceptButton = tabs.SelectedTab == singleTab ? compare : null; };

            inputs.Dock = DockStyle.Fill;
            inputs.Size = new Size(758, 360);
            layout.Controls.Add(inputs, 0, 1);
            inputs.Controls.Add(new Label { Text = "Job type", AutoSize = true, Location = new Point(0, 4), Font = new Font(Font, FontStyle.Bold) });
            jobType.DropDownStyle = ComboBoxStyle.DropDownList;
            jobType.Items.AddRange(new object[] { "Full Document", "EP - description, full specification and claims" });
            jobType.SetBounds(92, 0, 534, 29);
            jobType.Anchor = AnchorStyles.Left | AnchorStyles.Top | AnchorStyles.Right;
            jobType.SelectedIndexChanged += delegate
            {
                languageChecked.Visible = false;
                UpdateJobHint();
            };
            inputs.Controls.Add(jobType);
            reset.Text = "Reset";
            reset.Name = "resetSingle";
            reset.AccessibleName = "Reset single comparison";
            reset.AccessibleDescription = "Clear selected files and displayed results. Keep the job type, save folder and saved documents.";
            reset.SetBounds(642, 0, 92, 31);
            reset.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            BrandAppearance.StyleButton(reset, false);
            reset.Click += delegate { ResetComparison(); };
            inputs.Controls.Add(reset);
            jobHint.SetBounds(0, 33, 734, 45);
            jobHint.Anchor = AnchorStyles.Left | AnchorStyles.Top | AnchorStyles.Right;
            jobHint.ForeColor = BrandAppearance.Muted;
            jobHint.Font = BrandAppearance.Font(9F, FontStyle.Regular);
            inputs.Controls.Add(jobHint);
            jobType.SelectedIndex = 0;
            AddFileRow("1   Archive document", archive, archiveOcr, 80);
            AddFileRow("2   New document", revised, newOcr, 170);
            Label folderLabel = new Label { Text = "3   Save comparison to", AutoSize = true, Location = new Point(0, 261), Font = new Font(Font, FontStyle.Bold) };
            inputs.Controls.Add(folderLabel);
            destination.SetBounds(0, 287, 527, 29);
            destination.Anchor = AnchorStyles.Left | AnchorStyles.Top | AnchorStyles.Right;
            destination.Text = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "LocalDocumentCompare", "Results");
            inputs.Controls.Add(destination);
            Button desktopSave = MakeButton("Desktop", 92);
            desktopSave.Name = "saveDesktop";
            desktopSave.SetBounds(544, 286, 92, 31);
            desktopSave.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            desktopSave.Click += delegate { destination.Text = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory); };
            inputs.Controls.Add(desktopSave);
            Button folderBrowse = MakeButton("Browse…", 92);
            folderBrowse.Name = "browseOutput";
            folderBrowse.AccessibleName = "Browse for a comparison results folder";
            folderBrowse.SetBounds(642, 286, 92, 31);
            folderBrowse.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            folderBrowse.Click += delegate
            {
                using (FolderBrowserDialog picker = new FolderBrowserDialog { Description = "Choose a Desktop, local or shared-drive folder for comparison results", SelectedPath = destination.Text })
                    if (picker.ShowDialog(this) == DialogResult.OK) destination.Text = picker.SelectedPath;
            };
            inputs.Controls.Add(folderBrowse);
            inputs.Controls.Add(new Label { Name = "inputHint", Text = "Browse local or shared-drive locations, or drop files into the boxes above.\r\nPDF scan mode reads every page with local OCR and simplifies page layout.", AutoSize = false, Location = new Point(0, 324), Size = new Size(740, 30), Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right, ForeColor = BrandAppearance.Muted, Font = BrandAppearance.Font(9F, FontStyle.Regular) });

            FlowLayoutPanel actions = new FlowLayoutPanel { Dock = DockStyle.Fill, Padding = new Padding(0, 6, 0, 0), WrapContents = false };
            compare.Text = "Compare documents";
            compare.Size = new Size(206, 40);
            BrandAppearance.StyleButton(compare, true);
            compare.Font = new Font(Font, FontStyle.Bold);
            compare.Click += StartComparison;
            cancel.Text = "Cancel";
            cancel.Size = new Size(95, 40);
            BrandAppearance.StyleButton(cancel, false);
            cancel.Enabled = false;
            cancel.Click += delegate { CancelComparison(); };
            actions.Controls.Add(compare);
            actions.Controls.Add(cancel);
            languageChecked.Text = "Language checked — compare";
            languageChecked.Size = new Size(256, 40);
            BrandAppearance.StyleButton(languageChecked, false);
            languageChecked.Visible = false;
            languageChecked.Click += delegate { StartComparison(languageChecked, EventArgs.Empty); };
            actions.Controls.Add(languageChecked);
            layout.Controls.Add(actions, 0, 2);
            state.Text = "Ready — choose the archive and new document.";
            state.Dock = DockStyle.Fill;
            state.TextAlign = ContentAlignment.MiddleLeft;
            layout.Controls.Add(state, 0, 3);
            progress.Dock = DockStyle.Fill;
            progress.Visible = false;
            layout.Controls.Add(progress, 0, 4);
            metrics.Text = "Yellow insertions  ·  Tracked deletions and moves";
            metrics.Font = BrandAppearance.Font(13F, FontStyle.Bold);
            metrics.ForeColor = BrandAppearance.Purple;
            metrics.Dock = DockStyle.Fill;
            metrics.TextAlign = ContentAlignment.MiddleLeft;
            Panel metricPanel = new Panel { Dock = DockStyle.Fill };
            metricPanel.Controls.Add(metrics);
            ConfigureSectionMetrics();
            metricPanel.Controls.Add(sectionMetrics);
            layout.Controls.Add(metricPanel, 0, 5);
            details.Multiline = true;
            details.ReadOnly = true;
            details.BorderStyle = BorderStyle.None;
            details.BackColor = BrandAppearance.Surface;
            details.ForeColor = BrandAppearance.Muted;
            details.Dock = DockStyle.Fill;
            details.ScrollBars = ScrollBars.Vertical;
            details.Text = "Results: crossover percentage, total word count of the new file, and highlighted insertion word count.\r\nCrossover is the share of the new document’s words retained from the archive.\r\n\r\nComparison and OCR run on this PC. Source files stay unchanged. Shared drives and synced folders follow your organisation’s access and storage settings.";
            layout.Controls.Add(details, 0, 6);
            FlowLayoutPanel resultActions = new FlowLayoutPanel { Dock = DockStyle.Fill, Padding = new Padding(0, 8, 0, 0) };
            openDocument.Text = "Open comparison";
            openDocument.Size = new Size(160, 31);
            BrandAppearance.StyleButton(openDocument, false);
            openDocument.Enabled = false;
            openDocument.Click += delegate { OpenPath(lastOutput); };
            openFolder.Text = "Open results folder";
            openFolder.Size = new Size(170, 31);
            BrandAppearance.StyleButton(openFolder, false);
            openFolder.Enabled = false;
            openFolder.Click += delegate { OpenPath(Path.GetDirectoryName(lastOutput)); };
            resultActions.Controls.Add(openDocument);
            resultActions.Controls.Add(openFolder);
            layout.Controls.Add(resultActions, 0, 7);
            timer.Interval = 350;
            timer.Tick += PollWorker;
            FormClosing += OnClosing;
            AcceptButton = compare;
            inputs.Resize += delegate { LayoutInputs(); };
            LayoutInputs();
        }

        private void LayoutInputs()
        {
            int width = inputs.ClientSize.Width;
            if (width < 220) return;
            jobType.Width = Math.Max(1, width - jobType.Left - 108);
            reset.Left = width - reset.Width;
            jobHint.Width = width;
            archive.Width = revised.Width = width - 108;
            destination.Width = width - 207;
            foreach (string name in new[] { "browseArchive", "browseNew", "browseOutput" })
            {
                Control[] found = inputs.Controls.Find(name, false);
                if (found.Length == 1) found[0].Left = width - found[0].Width;
            }
            Control[] desktop = inputs.Controls.Find("saveDesktop", false);
            if (desktop.Length == 1) desktop[0].Left = width - 190;
            Control[] note = inputs.Controls.Find("inputHint", false);
            if (note.Length == 1) note[0].Width = width;
        }

        private static Button MakeButton(string text, int width)
        {
            var button = new Button { Text = text, Width = width };
            BrandAppearance.StyleButton(button, false);
            return button;
        }

        private Panel CreateHeading()
        {
            Panel heading = new Panel { Dock = DockStyle.Fill };
            heading.Controls.Add(new PictureBox { Image = brandLogo, SizeMode = PictureBoxSizeMode.Zoom,
                Location = new Point(0, 0), Size = new Size(162, 65), TabStop = false, AccessibleName = "RWS" });
            heading.Controls.Add(new Label { Text = "Document Compare", Font = BrandAppearance.Font(24F, FontStyle.Bold),
                AutoSize = true, Location = new Point(200, 4), ForeColor = BrandAppearance.Purple });
            heading.Controls.Add(new Label { Text = "Compare locally in Word. Keep the changes clear.",
                AutoSize = true, Location = new Point(202, 47), ForeColor = BrandAppearance.Muted,
                Font = BrandAppearance.Font(9F, FontStyle.Regular) });
            heading.Controls.Add(new Panel { BackColor = BrandAppearance.Pink, Location = new Point(0, 83), Size = new Size(72, 3) });
            heading.Controls.Add(new Panel { BackColor = BrandAppearance.Border, Location = new Point(72, 85), Size = new Size(682, 1),
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right });
            return heading;
        }

        private void UpdateModeAvailability()
        {
            singleTab.Enabled = !batch.IsBusy && !multiArchive.IsBusy;
            batchTab.Enabled = worker == null && !multiArchive.IsBusy;
            multiArchiveTab.Enabled = worker == null && !batch.IsBusy;
            reset.Enabled = worker == null && !batch.IsBusy && !multiArchive.IsBusy && !keepJobDirectory;
        }

        private void ResetComparison()
        {
            if (worker != null || batch.IsBusy || multiArchive.IsBusy || keepJobDirectory || IsDisposed || Disposing) return;
            timer.Stop();
            archive.Clear();
            revised.Clear();
            archiveOcr.Checked = newOcr.Checked = false;
            archiveOcr.Enabled = newOcr.Enabled = false;
            languageChecked.Visible = false;
            lastOutput = null;
            cancelling = false;
            started = default(DateTime);
            progress.Visible = false;
            progress.Style = ProgressBarStyle.Blocks;
            progress.Value = 0;
            sectionMetrics.Rows.Clear();
            sectionMetrics.Visible = false;
            metrics.Visible = true;
            metrics.Text = "Yellow insertions  ·  Tracked deletions and moves";
            state.Text = "Ready — choose the archive and new document.";
            details.Text = "Results: crossover percentage, total word count of the new file, and highlighted insertion word count.\r\nCrossover is the share of the new document’s words retained from the archive.\r\n\r\nComparison and OCR run on this PC. Source files stay unchanged. Shared drives and synced folders follow your organisation’s access and storage settings.";
            if (!String.IsNullOrEmpty(jobDirectory))
                details.AppendText("\r\n\r\nPrevious temporary files remain: " + jobDirectory);
            inputs.Enabled = compare.Enabled = true;
            cancel.Enabled = openDocument.Enabled = openFolder.Enabled = false;
            UpdateJobHint();
            UpdateModeAvailability();
            archive.Focus();
        }

        protected override void Dispose(bool disposing)
        {
            base.Dispose(disposing);
            if (disposing)
            {
                if (brandLogo != null) brandLogo.Dispose();
                if (brandIcon != null) brandIcon.Dispose();
                timer.Dispose();
            }
        }

        private void ConfigureSectionMetrics()
        {
            sectionMetrics.Dock = DockStyle.Fill;
            sectionMetrics.Visible = false;
            sectionMetrics.ReadOnly = true;
            sectionMetrics.AllowUserToAddRows = false;
            sectionMetrics.AllowUserToDeleteRows = false;
            sectionMetrics.AllowUserToResizeRows = false;
            sectionMetrics.AllowUserToResizeColumns = false;
            sectionMetrics.RowHeadersVisible = false;
            sectionMetrics.ScrollBars = ScrollBars.None;
            sectionMetrics.BorderStyle = BorderStyle.None;
            sectionMetrics.BackgroundColor = Color.White;
            sectionMetrics.GridColor = BrandAppearance.Border;
            sectionMetrics.CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal;
            sectionMetrics.ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.None;
            sectionMetrics.EnableHeadersVisualStyles = false;
            sectionMetrics.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
            sectionMetrics.ColumnHeadersHeight = 44;
            sectionMetrics.ColumnHeadersDefaultCellStyle.BackColor = BrandAppearance.Purple;
            sectionMetrics.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            sectionMetrics.ColumnHeadersDefaultCellStyle.Font = BrandAppearance.Font(9F, FontStyle.Bold);
            sectionMetrics.ColumnHeadersDefaultCellStyle.WrapMode = DataGridViewTriState.True;
            sectionMetrics.DefaultCellStyle.ForeColor = BrandAppearance.Text;
            sectionMetrics.DefaultCellStyle.BackColor = Color.White;
            sectionMetrics.DefaultCellStyle.SelectionForeColor = BrandAppearance.Text;
            sectionMetrics.DefaultCellStyle.SelectionBackColor = Color.White;
            sectionMetrics.RowTemplate.Height = 28;
            sectionMetrics.TabStop = false;
            sectionMetrics.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            AddMetricColumn("scope", "EP scope", 25F, false);
            AddMetricColumn("crossover", "Crossover", 17F, true);
            AddMetricColumn("newWords", "New file word count", 29F, true);
            AddMetricColumn("insertions", "Highlighted insertion word count", 29F, true);
        }

        private void AddMetricColumn(string name, string header, float weight, bool numeric)
        {
            DataGridViewTextBoxColumn column = new DataGridViewTextBoxColumn
            {
                Name = name, HeaderText = header, FillWeight = weight,
                SortMode = DataGridViewColumnSortMode.NotSortable
            };
            column.DefaultCellStyle.Alignment = numeric ? DataGridViewContentAlignment.MiddleRight : DataGridViewContentAlignment.MiddleLeft;
            column.HeaderCell.Style.Alignment = numeric ? DataGridViewContentAlignment.MiddleRight : DataGridViewContentAlignment.MiddleLeft;
            sectionMetrics.Columns.Add(column);
        }

        private void ShowResultMetrics(ComparisonResult result)
        {
            bool hasSections = result.Sections != null && result.Sections.Length > 0;
            sectionMetrics.Visible = hasSections;
            metrics.Visible = !hasSections;
            sectionMetrics.Rows.Clear();
            if (hasSections)
            {
                foreach (ComparisonSectionResult section in result.Sections)
                {
                    sectionMetrics.Rows.Add(section.Scope,
                        section.CrossoverPercent.HasValue ? section.CrossoverPercent.Value.ToString("0.0") + "%" : "Unavailable",
                        section.NewWords.ToString("N0"), section.InsertedWords.ToString("N0"));
                }
                sectionMetrics.ClearSelection();
            }
            else
            {
                metrics.Text = "Crossover: " + (result.CrossoverPercent.HasValue ? result.CrossoverPercent.Value.ToString("0.0") + "%" : "Unavailable") + "\r\nTotal word count of new file: " + result.NewWords.ToString("N0") + "\r\nHighlighted insertion word count: " + result.InsertedWords.ToString("N0");
            }
        }

        private void AddFileRow(string title, TextBox box, CheckBox ocr, int top)
        {
            inputs.Controls.Add(new Label { Text = title, AutoSize = true, Location = new Point(0, top), Font = new Font(Font, FontStyle.Bold) });
            box.SetBounds(0, top + 27, 625, 29);
            box.Anchor = AnchorStyles.Left | AnchorStyles.Top | AnchorStyles.Right;
            box.AllowDrop = true;
            box.DragEnter += delegate(object sender, DragEventArgs e) { if (e.Data.GetDataPresent(DataFormats.FileDrop)) e.Effect = DragDropEffects.Copy; };
            box.DragDrop += delegate(object sender, DragEventArgs e)
            {
                string[] paths = e.Data.GetData(DataFormats.FileDrop) as string[];
                if (paths != null && paths.Length == 1) box.Text = paths[0];
            };
            box.TextChanged += delegate
            {
                languageChecked.Visible = false;
                ocr.Enabled = box.Text.Trim().Trim('"').EndsWith(".pdf", StringComparison.OrdinalIgnoreCase);
                if (!ocr.Enabled) ocr.Checked = false;
                if (Object.ReferenceEquals(box, revised)) UpdateJobHint();
            };
            ocr.CheckedChanged += delegate { languageChecked.Visible = false; };
            inputs.Controls.Add(box);
            Button browse = MakeButton("Browse…", 92);
            BrandAppearance.StyleButton(browse, true);
            browse.Name = Object.ReferenceEquals(box, archive) ? "browseArchive" : "browseNew";
            browse.AccessibleName = "Browse for the " + (Object.ReferenceEquals(box, archive) ? "archive" : "new") + " document";
            browse.AccessibleDescription = "Choose a Word or PDF document from a local or shared-drive location.";
            browse.SetBounds(642, top + 26, 92, 31);
            browse.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            browse.Click += delegate
            {
                using (OpenFileDialog picker = new OpenFileDialog { Title = "Choose the " + title.Substring(4).ToLowerInvariant() + " — local or shared drive", Filter = "Word and PDF documents|*.docx;*.doc;*.pdf", CheckFileExists = true, Multiselect = false, RestoreDirectory = true })
                    if (picker.ShowDialog(this) == DialogResult.OK) box.Text = picker.FileName;
            };
            inputs.Controls.Add(browse);
            ocr.Text = "PDF has scanned pages — use local OCR";
            ocr.AutoSize = true;
            ocr.Location = new Point(0, top + 60);
            ocr.Enabled = false;
            ocr.Font = BrandAppearance.Font(9F, FontStyle.Regular);
            inputs.Controls.Add(ocr);
        }

        private void UpdateJobHint()
        {
            bool ep = jobType.SelectedIndex == 1;
            bool newPdf = revised.Text.Trim().Trim('"').EndsWith(".pdf", StringComparison.OrdinalIgnoreCase);
            jobHint.Text = ep
                ? (newPdf ? "The new file is a PDF: choose Full Document for a reliable whole-document comparison." : "For EP PDFs or Word files converted from PDF, choose Full Document.") + "\r\nEP needs clear description and claims sections; converted sections can be unreliable."
                : "Full Document compares all text in both documents.\r\nRecommended for EP PDFs and Word files converted from PDF.";
            jobHint.ForeColor = ep && newPdf ? BrandAppearance.Pink : BrandAppearance.Muted;
        }

        private void StartComparison(object sender, EventArgs e)
        {
            if (worker != null || batch.IsBusy || multiArchive.IsBusy) return;
            try
            {
                bool languageConfirmed = Object.ReferenceEquals(sender, languageChecked);
                string archivePath = ValidateInput(archive.Text, "archive");
                string newPath = ValidateInput(revised.Text, "new");
                if (String.Equals(archivePath, newPath, StringComparison.OrdinalIgnoreCase)) throw new InvalidOperationException("Choose two different files for the archive and new documents.");
                string outputFolder = Path.GetFullPath(destination.Text.Trim().Trim('"'));
                ComparisonEngine.RequireOutputPath(outputFolder);
                string outputPath = Path.Combine(outputFolder, Path.GetFileNameWithoutExtension(newPath) + " vs " + Path.GetFileNameWithoutExtension(archivePath) + ".docx");
                if (outputPath.Length > 240) throw new InvalidOperationException("The result filename is too long. Use shorter source filenames or a shorter destination folder.");
                if (File.Exists(outputPath) || File.Exists(Path.ChangeExtension(outputPath, ".txt"))) throw new InvalidOperationException("A comparison with this name already exists. Choose another destination folder, or move the earlier result first.");
                if (Type.GetTypeFromProgID("Word.Application") == null) throw new InvalidOperationException("Install and activate desktop Microsoft Word before using this tool.");
                Directory.CreateDirectory(outputFolder);
                jobDirectory = Path.Combine(Path.GetTempPath(), "LocalDocumentCompare", Guid.NewGuid().ToString("N"));
                ComparisonEngine.RequireLocalPath(jobDirectory);
                Directory.CreateDirectory(jobDirectory);
                ComparisonRequest request = new ComparisonRequest { ArchivePath = archivePath, NewPath = newPath, OutputPath = outputPath, OcrArchive = archiveOcr.Checked, OcrNew = newOcr.Checked, WorkDirectory = jobDirectory, JobType = jobType.SelectedIndex == 1 ? "EP" : "WF", MatchingLanguageConfirmed = languageConfirmed };
                string requestPath = Path.Combine(jobDirectory, "request.json");
                JsonFile.Write(requestPath, request);
                ProcessStartInfo start = new ProcessStartInfo(System.Reflection.Assembly.GetExecutingAssembly().Location, "--worker " + Program.Quote(requestPath) + " " + Program.Quote(Path.Combine(jobDirectory, "result.json")))
                {
                    UseShellExecute = false, CreateNoWindow = true, WindowStyle = ProcessWindowStyle.Hidden,
                    WorkingDirectory = AppDomain.CurrentDomain.BaseDirectory
                };
                worker = Process.Start(start);
                started = DateTime.UtcNow;
                cancelling = false;
                keepJobDirectory = false;
                inputs.Enabled = false;
                compare.Enabled = false;
                cancel.Enabled = true;
                languageChecked.Visible = false;
                openDocument.Enabled = false;
                openFolder.Enabled = false;
                progress.Style = ProgressBarStyle.Marquee;
                progress.Visible = true;
                state.Text = "Starting a private Word comparison session…";
                sectionMetrics.Visible = false;
                metrics.Visible = true;
                metrics.Text = "Comparing documents";
                details.Text = "The source documents are opened as temporary copies. Larger PDFs and scanned pages can take several minutes.";
                timer.Start();
                UpdateModeAvailability();
            }
            catch (Exception ex)
            {
                if (worker == null) CleanupJob();
                MessageBox.Show(this, ex.Message, "Cannot start comparison", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private static string ValidateInput(string value, string label)
        {
            if (String.IsNullOrWhiteSpace(value)) throw new InvalidOperationException("Choose the " + label + " document.");
            string path = Path.GetFullPath(value.Trim().Trim('"'));
            ComparisonEngine.RequireSourcePath(path);
            string extension = Path.GetExtension(path).ToLowerInvariant();
            if (extension != ".doc" && extension != ".docx" && extension != ".pdf") throw new InvalidOperationException("The " + label + " file must be a .doc, .docx or .pdf document.");
            if (!File.Exists(path)) throw new InvalidOperationException("The " + label + " document could not be found or read. Check file access and any shared-drive connection.");
            return path;
        }

        private void PollWorker(object sender, EventArgs e)
        {
            if (worker == null) return;
            try
            {
                string statusPath = Path.Combine(jobDirectory, "status.txt");
                if (File.Exists(statusPath))
                {
                    try { state.Text = File.ReadAllText(statusPath, Encoding.UTF8); } catch (IOException) { }
                }
                if (!worker.HasExited)
                {
                    if ((DateTime.UtcNow - started).TotalMinutes >= 20)
                    {
                        CancelComparison();
                        details.Text = "The comparison exceeded 20 minutes and was stopped. Try a smaller document, or open both files in Word to check for password or conversion prompts." + Environment.NewLine + details.Text;
                    }
                    return;
                }
                FinishComparison();
            }
            catch (Exception)
            {
                CancelComparison();
                details.Text = "The comparison process stopped unexpectedly. Check the files in Word, then try again.\r\n" + details.Text;
            }
        }

        private void FinishComparison()
        {
            timer.Stop();
            ComparisonResult result = null;
            string resultPath = Path.Combine(jobDirectory, "result.json");
            if (File.Exists(resultPath)) result = JsonFile.Read<ComparisonResult>(resultPath);
            StopOwnedChildren();
            worker.Dispose();
            worker = null;
            UpdateModeAvailability();
            progress.Visible = false;
            inputs.Enabled = true;
            compare.Enabled = true;
            cancel.Enabled = false;
            languageChecked.Visible = false;
            if (result != null && result.Success)
            {
                lastOutput = result.OutputPath;
                state.Text = "Comparison saved.";
                ShowResultMetrics(result);
                details.Text = result.OutputPath + "\r\nA summary is saved alongside the Word document.";
                if (result.Warnings != null && result.Warnings.Length > 0) details.AppendText("\r\n\r\n" + String.Join("\r\n", result.Warnings));
                openDocument.Enabled = true;
                openFolder.Enabled = true;
            }
            else
            {
                state.Text = cancelling ? "Comparison cancelled." : result != null && result.LanguageMismatch != null ? "Comparison stopped — source languages do not match." : result != null && result.LanguageReviewRequired ? "Language check required — comparison stopped." : "Comparison could not finish.";
                sectionMetrics.Visible = false;
                metrics.Visible = true;
                metrics.Text = "No completed comparison";
                details.Text = cancelling ? "The comparison session was stopped. The source documents are unchanged." : (result == null ? "Word may have been closed or a conversion prompt may have interrupted the comparison. Check the source documents in Word and try again." : result.Error);
                languageChecked.Visible = !cancelling && result != null && result.LanguageReviewRequired;
            }
            string leftover = CleanupJob();
            if (leftover != null) details.AppendText("\r\n\r\nSome temporary files could not be removed. Close Word and delete this folder: " + leftover);
            if (!cancelling && !IsDisposed && !Disposing) LanguageMismatchNotice.ShowSingle(this, result);
        }

        private bool CancelComparison()
        {
            if (worker == null) return true;
            cancelling = true;
            timer.Stop();
            try
            {
                File.WriteAllText(Path.Combine(jobDirectory, "cancel.flag"), "cancel");
                if (!worker.HasExited && !worker.WaitForExit(1200)) { worker.Kill(); worker.WaitForExit(3000); }
            }
            catch { }
            try
            {
                if (!worker.HasExited)
                {
                    state.Text = "The comparison process could not be stopped.";
                    details.Text = "Temporary files remain in " + jobDirectory + ". Try Cancel again before closing.";
                    timer.Start();
                    return false;
                }
            }
            catch { return false; }
            StopOwnedChildren();
            FinishComparison();
            return true;
        }

        private void StopOwnedChildren()
        {
            if (String.IsNullOrEmpty(jobDirectory)) return;
            if (!StopRecordedProcess("ocr", "powershell")) keepJobDirectory = true;
            if (!StopRecordedProcess("word", "WINWORD")) keepJobDirectory = true;
        }

        private bool StopRecordedProcess(string name, string expectedName)
        {
            try
            {
                int pid;
                long ticks;
                string recordPath = Path.Combine(jobDirectory, name + ".process");
                if (File.Exists(recordPath))
                {
                    OwnedProcess record = JsonFile.Read<OwnedProcess>(recordPath);
                    pid = record.Id;
                    ticks = record.StartTimeUtcTicks;
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
            catch (ArgumentException) { return true; } // The recorded process already exited.
            catch { return false; }
        }

        private string CleanupJob()
        {
            if (String.IsNullOrEmpty(jobDirectory)) return null;
            string path = jobDirectory;
            if (keepJobDirectory) return path;
            try
            {
                string allowedRoot = Path.GetFullPath(Path.Combine(Path.GetTempPath(), "LocalDocumentCompare")) + Path.DirectorySeparatorChar;
                string resolved = Path.GetFullPath(path);
                if (!resolved.StartsWith(allowedRoot, StringComparison.OrdinalIgnoreCase)) return path;
                if (Directory.Exists(resolved)) Directory.Delete(resolved, true);
                jobDirectory = null;
                return null;
            }
            catch { return path; }
        }

        private void OpenPath(string path)
        {
            try { if (!String.IsNullOrEmpty(path)) Process.Start(new ProcessStartInfo(path) { UseShellExecute = true }); }
            catch { MessageBox.Show(this, "The saved result could not be opened. Browse to the destination folder to open it.", "Document Compare"); }
        }

        private void OnClosing(object sender, FormClosingEventArgs e)
        {
            if (multiArchive.IsBusy)
            {
                if (MessageBox.Show(this, "Stop the archive search and close? Completed comparisons will be kept.", "Document Compare", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes)
                { e.Cancel = true; return; }
                if (!multiArchive.CancelAndClose(false)) { e.Cancel = true; return; }
            }
            if (batch.IsBusy)
            {
                if (MessageBox.Show(this, "Stop the batch and close? Completed comparisons will be kept.", "Document Compare", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes)
                { e.Cancel = true; return; }
                if (!batch.CancelAndClose(false)) { e.Cancel = true; return; }
            }
            if (worker == null) return;
            if (MessageBox.Show(this, "Stop the current comparison and close?", "Document Compare", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes) { e.Cancel = true; return; }
            if (!CancelComparison()) { e.Cancel = true; return; }
            if (jobDirectory != null) MessageBox.Show(this, "Some temporary files remain. Close Word and remove this folder when it is no longer in use:\r\n" + jobDirectory, "Temporary files", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private sealed class OwnedProcess
        {
            public OwnedProcess() { }
            public int Id { get; set; }
            public long StartTimeUtcTicks { get; set; }
        }
    }
}
