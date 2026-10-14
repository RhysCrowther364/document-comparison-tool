using System;
using System.Drawing;
using System.Drawing.Text;
using System.IO;
using System.Reflection;
using System.Windows.Forms;

namespace LocalDocumentCompare
{
    // Colours and original header logo from the RWS May 2025 signage guide.
    // The application icon uses a separate document-comparison design.
    // Articulat CF is used only when installed; no fonts or online assets are loaded.
    internal static class BrandAppearance
    {
        internal static readonly Color Purple = Color.FromArgb(62, 1, 111);
        internal static readonly Color Pink = Color.FromArgb(230, 0, 84);
        internal static readonly Color Text = Color.FromArgb(24, 24, 24);
        internal static readonly Color Muted = Color.FromArgb(85, 85, 85);
        internal static readonly Color Surface = Color.FromArgb(246, 246, 246);
        internal static readonly Color Border = Color.FromArgb(218, 218, 218);
        internal static readonly string FontName = FindFont();

        private static string FindFont()
        {
            using (var fonts = new InstalledFontCollection())
                foreach (FontFamily family in fonts.Families)
                    if (String.Equals(family.Name, "Articulat CF", StringComparison.OrdinalIgnoreCase)) return family.Name;
            return "Arial";
        }

        internal static Font Font(float size, FontStyle style)
        {
            try { return new Font(FontName, size, style, GraphicsUnit.Point); }
            catch (ArgumentException) { return new Font("Arial", size, style, GraphicsUnit.Point); }
        }

        internal static Image LoadLogo()
        {
            using (Stream stream = Assembly.GetExecutingAssembly().GetManifestResourceStream("LocalDocumentCompare.RwsLogo.png"))
            {
                if (stream == null) throw new InvalidOperationException("The application logo is missing. Reinstall the complete package.");
                using (Image image = Image.FromStream(stream)) return new Bitmap(image);
            }
        }

        internal static Icon LoadIcon()
        {
            using (Stream stream = Assembly.GetExecutingAssembly().GetManifestResourceStream("LocalDocumentCompare.ComparisonIcon.ico"))
            {
                if (stream == null) throw new InvalidOperationException("The application icon is missing. Reinstall the complete package.");
                using (var icon = new Icon(stream)) return (Icon)icon.Clone();
            }
        }

        internal static void StyleButton(Button button, bool primary)
        {
            button.FlatStyle = FlatStyle.Flat;
            button.UseVisualStyleBackColor = false;
            button.BackColor = primary ? Purple : Color.White;
            button.ForeColor = primary ? Color.White : Purple;
            button.FlatAppearance.BorderSize = primary ? 0 : 1;
            button.FlatAppearance.BorderColor = Purple;
            button.FlatAppearance.MouseOverBackColor = primary ? Color.FromArgb(82, 20, 132) : Surface;
            button.FlatAppearance.MouseDownBackColor = primary ? Color.FromArgb(44, 0, 80) : Color.FromArgb(235, 235, 235);
            button.Cursor = Cursors.Hand;
            button.Padding = new Padding(4, 0, 4, 0);
        }
    }
}
