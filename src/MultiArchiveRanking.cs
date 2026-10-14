using System;
using System.Collections.Generic;
using System.Linq;

namespace LocalDocumentCompare
{
    public static class MultiArchiveRanking
    {
        public static double? Score(BatchRecord record, string jobType)
        {
            if (record == null || record.Pair == null || record.State != "Completed" || record.Result == null ||
                !record.Result.Success || record.Result.LanguageReviewRequired) return null;
            ComparisonResult result = record.Result;
            int words = result.NewWords, inserted = result.InsertedWords;
            double? score = result.CrossoverPercent;
            if (String.Equals(jobType, "EP", StringComparison.OrdinalIgnoreCase))
            {
                string[] required = { "Description", "Full specification", "Claims" };
                if (result.Sections == null || required.Any(scope => result.Sections.Count(s => s != null &&
                    String.Equals(s.Scope, scope, StringComparison.OrdinalIgnoreCase)) != 1)) return null;
                ComparisonSectionResult full = result.Sections.Single(s => s != null &&
                    String.Equals(s.Scope, "Full specification", StringComparison.OrdinalIgnoreCase));
                words = full.NewWords; inserted = full.InsertedWords; score = full.CrossoverPercent;
            }
            // An unavailable percentage must stay unranked, even when counts
            // are present: the comparison engine may have withheld it for review.
            if (!score.HasValue || Double.IsNaN(score.Value) || Double.IsInfinity(score.Value) ||
                score.Value < 0 || score.Value > 100 || words <= 0 || inserted < 0 || inserted > words ||
                Math.Abs(score.Value - 100.0 * (words - inserted) / words) > 0.000001) return null;
            return score;
        }

        public static List<BatchRecord> Ordered(IEnumerable<BatchRecord> records, string jobType)
        {
            return records.Where(r => r != null && r.Pair != null)
                .OrderByDescending(r => Score(r, jobType).HasValue)
                .ThenByDescending(r => Score(r, jobType) ?? -1)
                .ThenBy(r => r.Pair.Number).ToList();
        }

        public static Dictionary<int, int> Ranks(IEnumerable<BatchRecord> records, string jobType)
        {
            var ranks = new Dictionary<int, int>();
            double? previous = null;
            int index = 0, rank = 0;
            foreach (BatchRecord record in Ordered(records, jobType))
            {
                double? score = Score(record, jobType);
                if (!score.HasValue) continue;
                index++;
                if (!previous.HasValue || score.Value != previous.Value) rank = index;
                ranks.Add(record.Pair.Number, rank);
                previous = score;
            }
            return ranks;
        }
    }
}
