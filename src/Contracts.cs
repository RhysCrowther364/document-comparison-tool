using System;
using System.IO;
using System.Text;
using System.Web.Script.Serialization;

namespace LocalDocumentCompare
{
    public sealed class ComparisonRequest
    {
        public string ArchivePath { get; set; }
        public string NewPath { get; set; }
        public string OutputPath { get; set; }
        public bool OcrArchive { get; set; }
        public bool OcrNew { get; set; }
        public string WorkDirectory { get; set; }
        public string JobType { get; set; }
        public bool MatchingLanguageConfirmed { get; set; }
    }

    public sealed class ComparisonResult
    {
        public bool Success { get; set; }
        public string Error { get; set; }
        public string OutputPath { get; set; }
        public string SummaryPath { get; set; }
        public int NewWords { get; set; }
        public int InsertedWords { get; set; }
        public double? CrossoverPercent { get; set; }
        public string[] Warnings { get; set; }
        public string JobType { get; set; }
        public ComparisonSectionResult[] Sections { get; set; }
        public bool LanguageReviewRequired { get; set; }
        public SourceLanguageMismatch LanguageMismatch { get; set; }
    }

    public sealed class SourceLanguageMismatch
    {
        public string Scope { get; set; }
        public string ArchiveLanguage { get; set; }
        public string NewLanguage { get; set; }
    }

    public sealed class ComparisonSectionResult
    {
        public string Scope { get; set; }
        public int NewWords { get; set; }
        public int InsertedWords { get; set; }
        public double? CrossoverPercent { get; set; }
    }

    public static class JsonFile
    {
        public static T Read<T>(string path)
        {
            return new JavaScriptSerializer().Deserialize<T>(File.ReadAllText(path, Encoding.UTF8));
        }
        public static void Write(string path, object value)
        {
            File.WriteAllText(path, new JavaScriptSerializer().Serialize(value), new UTF8Encoding(false));
        }
    }
}
