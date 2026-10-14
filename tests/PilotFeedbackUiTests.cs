using System;
using System.Collections.Generic;
using System.Drawing;
using System.Reflection;
using System.Windows.Forms;
using LocalDocumentCompare;

internal static class PilotFeedbackUiTests
{
    private static int checks;

    [STAThread]
    public static int Main()
    {
        try
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            using (MainForm form = new MainForm())
            {
                form.Opacity = 0; form.ShowInTaskbar = false;
                form.Show(); Application.DoEvents();
                ComboBox jobType = Field<ComboBox>(form, "jobType");
                TextBox revised = Field<TextBox>(form, "revised");
                Label hint = Field<Label>(form, "jobHint");
                Panel inputs = Field<Panel>(form, "inputs");
                Button archiveBrowse = Named<Button>(inputs, "browseArchive");
                Button newBrowse = Named<Button>(inputs, "browseNew");
                Button outputBrowse = Named<Button>(inputs, "browseOutput");
                Equal(true, archiveBrowse.Visible && newBrowse.Visible && outputBrowse.Visible, "single Browse buttons are visible");
                Equal(true, archiveBrowse.Enabled && newBrowse.Enabled && outputBrowse.Enabled, "single Browse buttons are available when idle");
                Equal(BrandAppearance.Purple, archiveBrowse.BackColor, "archive Browse uses a prominent filled brand button");
                Equal(BrandAppearance.Purple, newBrowse.BackColor, "new Browse uses a prominent filled brand button");
                Equal(true, newBrowse.AccessibleDescription.Contains("shared-drive"), "file browser identifies shared-drive support");
                Equal(true, hint.Text.Contains("Word files converted from PDF"), "WF explains converted EP file guidance");

                jobType.SelectedIndex = 1;
                Equal(true, hint.Text.Contains("choose Full Document") && hint.Text.Contains("description and claims"), "EP explains when Full Document is preferable");
                revised.Text = "\"Z:\\Pilot\\New.PDF\"";
                Equal(true, hint.Text.Contains("The new file is a PDF"), "quoted mapped-drive PDF gives a specific EP warning");
                Equal(BrandAppearance.Pink, hint.ForeColor, "EP PDF guidance stands out in brand pink");
                Equal(1, jobType.SelectedIndex, "PDF guidance leaves EP selected");
                Equal(true, Field<CheckBox>(form, "newOcr").Enabled, "quoted PDF still allows local OCR");
                revised.Text = @"\\server\team\Converted.docx";
                Equal(true, hint.Text.Contains("Word files converted from PDF"), "DOCX keeps guidance because conversion origin cannot be inferred");
                Equal(BrandAppearance.Muted, hint.ForeColor, "ordinary DOCX has general guidance colour");
                Equal(1, jobType.SelectedIndex, "DOCX does not change selected job type");

                inputs.Enabled = false;
                Equal(false, archiveBrowse.Enabled || newBrowse.Enabled || outputBrowse.Enabled, "single Browse buttons share the job input lock");
                inputs.Enabled = true;
                Equal(true, archiveBrowse.Enabled && newBrowse.Enabled && outputBrowse.Enabled, "single Browse buttons recover with the inputs");

                form.Size = form.MinimumSize; Application.DoEvents();
                InBounds(archiveBrowse, inputs, "archive Browse fits the minimum-width app");
                InBounds(newBrowse, inputs, "new Browse fits the minimum-width app");
                InBounds(outputBrowse, inputs, "save Browse fits the minimum-width app");
                Equal(true, hint.Bottom <= revised.Top - 70, "single guidance does not overlap file rows");
                Equal(true, Field<System.Diagnostics.Process>(form, "worker") == null, "single checks never start Word");
                form.Close();
            }

            using (Form host = new Form { Opacity = 0, ShowInTaskbar = false, Width = 760, Height = 950 })
            using (BatchComparisonControl control = new BatchComparisonControl { Dock = DockStyle.Fill })
            {
                host.Controls.Add(control); host.Show(); Application.DoEvents();
                ComboBox jobType = Field<ComboBox>(control, "jobType");
                Label hint = Field<Label>(control, "jobHint");
                Panel configuration = Field<Panel>(control, "configuration");
                Button browse = Named<Button>(configuration, "browseFolder");
                Button outputBrowse = Named<Button>(configuration, "browseOutput");
                List<BatchInput> files = Field<List<BatchInput>>(control, "files");
                Equal(true, browse.Visible && outputBrowse.Visible, "batch Browse buttons are visible");
                Equal(BrandAppearance.Purple, browse.BackColor, "batch input Browse matches prominent single-file buttons");
                Equal(true, Field<TextBox>(control, "details").Text.Contains("shared-drive"), "batch describes browsing shared folders");
                Equal(true, hint.Text.Contains("Word files converted from PDF"), "batch WF includes converted EP guidance");
                jobType.SelectedIndex = 1;
                files.Add(new BatchInput { Path = @"\\server\team\01.docx" });
                files.Add(new BatchInput { Path = @"Z:\Pilot\02.pdf" });
                Invoke(control, "RefreshFiles", 0);
                Equal(true, hint.Text.Contains("New files include PDFs"), "batch later PDF shows specific EP advice");
                Equal(BrandAppearance.Pink, hint.ForeColor, "batch new PDF guidance is highlighted");
                Equal(1, jobType.SelectedIndex, "batch PDF advice keeps EP selected");
                BatchInput first = files[0]; files[0] = files[1]; files[1] = first;
                Invoke(control, "RefreshFiles", 0);
                Equal(false, hint.Text.Contains("New files include PDFs"), "a first-only archive PDF is not labelled a new PDF");
                Equal(true, hint.Text.Contains("Word files converted from PDF"), "batch generic advice remains after reorder");
                Invoke(control, "SetBusy", true);
                Equal(false, browse.Enabled || outputBrowse.Enabled, "batch browsers lock while running");
                Set(control, "reviewPaused", true); Invoke(control, "UpdateActions");
                Equal(false, browse.Enabled || outputBrowse.Enabled, "batch browsers remain locked during language review");
                Invoke(control, "SetBusy", false);
                Equal(true, browse.Enabled && outputBrowse.Enabled, "batch browsers unlock after completion");
                InBounds(browse, configuration, "batch input Browse fits the minimum-width host");
                InBounds(outputBrowse, configuration, "batch save Browse fits the minimum-width host");
                Equal(true, hint.Bottom < Field<TextBox>(control, "folderPath").Top, "batch guidance does not overlap the folder row");
                Equal(true, Field<System.Diagnostics.Process>(control, "worker") == null, "batch checks never start Word");
                host.Close();
            }
            Console.WriteLine("Pilot feedback UI: " + checks + " checks passed; no Word comparisons started.");
            return 0;
        }
        catch (Exception error) { Console.Error.WriteLine(error); return 1; }
    }

    private static T Field<T>(object instance, string name) { return (T)instance.GetType().GetField(name, BindingFlags.Instance | BindingFlags.NonPublic).GetValue(instance); }
    private static void Set(object instance, string name, object value) { instance.GetType().GetField(name, BindingFlags.Instance | BindingFlags.NonPublic).SetValue(instance, value); }
    private static object Invoke(object instance, string name, params object[] arguments) { return instance.GetType().GetMethod(name, BindingFlags.Instance | BindingFlags.NonPublic).Invoke(instance, arguments); }
    private static T Named<T>(Control parent, string name) where T : Control
    {
        Control[] found = parent.Controls.Find(name, true);
        if (found.Length != 1) throw new Exception("Expected one " + name + " control; found " + found.Length);
        return (T)found[0];
    }
    private static void InBounds(Control button, Control parent, string name)
    {
        if (button.Left < 0 || button.Top < 0 || button.Right > parent.ClientSize.Width || button.Bottom > parent.ClientSize.Height)
            throw new Exception(name + ": button " + button.Bounds + "; parent " + parent.ClientSize);
        checks++;
    }
    private static void Equal<T>(T expected, T actual, string name)
    {
        if (!EqualityComparer<T>.Default.Equals(expected, actual)) throw new Exception(name + ": expected " + expected + ", got " + actual);
        checks++;
    }
}
