using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace LocalDocumentCompare
{
    internal static class LanguageMismatchNotice
    {
        internal const string SingleTitle = "Comparison stopped — language mismatch";
        internal const string BatchTitle = "Batch — source language mismatches";

        internal static string SingleMessage(ComparisonResult result)
        {
            if (!IsMismatch(result)) return null;
            SourceLanguageMismatch mismatch = result.LanguageMismatch;
            return "The comparison stopped because the source languages were detected as different.\r\n\r\n" +
                "Section: " + Scope(mismatch.Scope) + "\r\n" +
                "Archive: " + mismatch.ArchiveLanguage + "\r\n" +
                "New: " + mismatch.NewLanguage + "\r\n\r\n" +
                "Choose archive and new documents in the same source language. If both are already in the same language, check the detected languages and selected sections.\r\n\r\n" +
                "No comparison document has been created.";
        }

        internal static string BatchMessage(IList<BatchRecord> records)
        {
            if (records == null) return null;
            List<BatchRecord> mismatches = records.Where(r => r != null && r.Pair != null &&
                r.State == "Failed" && IsMismatch(r.Result)).ToList();
            if (mismatches.Count == 0) return null;
            StringBuilder text = new StringBuilder();
            text.Append(mismatches.Count == 1 ? "1 comparison could not be completed" : mismatches.Count + " comparisons could not be completed");
            text.Append(" because the source languages were detected as different.\r\n\r\n");
            foreach (BatchRecord record in mismatches.Take(5))
            {
                SourceLanguageMismatch mismatch = record.Result.LanguageMismatch;
                text.Append("Pair ").Append(record.Pair.Number).Append(": ")
                    .Append(FileName(record.Pair.New.Path)).Append(" vs ")
                    .Append(FileName(record.Pair.Archive.Path)).Append("\r\n")
                    .Append(Scope(mismatch.Scope)).Append(" — archive: ").Append(mismatch.ArchiveLanguage)
                    .Append("; new: ").Append(mismatch.NewLanguage).Append("\r\n\r\n");
            }
            if (mismatches.Count > 5) text.Append("Further affected pairs are listed in the batch results table.\r\n\r\n");
            text.Append("No comparison document was created for these pairs. Other completed comparisons are saved.\r\n\r\n")
                .Append("Choose matching-language source versions and compare the affected pairs again. If the detected languages appear incorrect, check the selected text and sections. The batch results table and CSV retain each failure reason.");
            return text.ToString();
        }

        internal static void ShowSingle(IWin32Window owner, ComparisonResult result)
        {
            string message = SingleMessage(result);
            if (message != null) MessageBox.Show(owner, message, SingleTitle, MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }

        internal static void ShowBatch(IWin32Window owner, IList<BatchRecord> records)
        {
            string message = BatchMessage(records);
            if (message != null) MessageBox.Show(owner, message, BatchTitle, MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }

        internal static void ShowMultiArchive(IWin32Window owner, IList<BatchRecord> records)
        {
            string message = BatchMessage(records);
            if (message != null) MessageBox.Show(owner, message.Replace("batch results table", "Multi Archive results table"),
                "Multi Archive — source language mismatches", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }

        private static bool IsMismatch(ComparisonResult result)
        {
            return result != null && !result.Success && !result.LanguageReviewRequired && result.LanguageMismatch != null;
        }

        private static string Scope(string value)
        {
            return String.Equals(value, "Document", StringComparison.OrdinalIgnoreCase) ? "Full Document" : value;
        }

        private static string FileName(string path)
        {
            string name = Path.GetFileName(path) ?? "";
            return name.Length > 80 ? name.Substring(0, 77) + "..." : name;
        }
    }
}
