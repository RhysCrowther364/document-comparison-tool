using System;
using System.Web.Script.Serialization;
using LocalDocumentCompare;

internal static class DiscountTests
{
    private static int checks;
    private static void Check(int total, int inserted, double? crossover, int? expected)
    {
        var result = new ComparisonResult { NewWords = total, InsertedWords = inserted, CrossoverPercent = crossover };
        if (result.DiscountPercent != expected) throw new Exception("Unexpected discount for " + inserted + "/" + total);
        var serializer = new JavaScriptSerializer();
        var loaded = serializer.Deserialize<ComparisonResult>(serializer.Serialize(result));
        if (loaded.DiscountPercent != expected) throw new Exception("Discount changed across worker/UI JSON transfer");
        checks++;
    }
    public static int Main()
    {
        Check(1000, 250, 75, 20);
        Check(1000, 0, 100, 45);
        Check(1000, 1000, 0, 0);
        Check(1000, 450, 55, 0);
        Check(1000, 425, 57.5, 3);
        Check(1000, 475, 52.5, 0);
        Check(1000, 451, 54.9, 0);
        Check(1000, 456, 54.4, 0);
        Check(1000, 424, 57.6, 3);
        Check(1000, 426, 57.4, 2);
        Check(0, 0, null, null);
        Check(1000, 250, null, null);
        Console.WriteLine("PASS " + checks + " discount cases: formula, zero minimum, rounding, empty/unreconciled text, and worker/UI serialization.");
        return 0;
    }
}
