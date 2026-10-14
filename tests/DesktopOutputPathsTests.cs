using System;
using System.Collections.Generic;
using System.IO;
using LocalDocumentCompare;

public static class DesktopOutputPathsTests
{
    private static int checks;

    public static int Main()
    {
        string[] names = { "OneDrive", "OneDriveConsumer", "OneDriveCommercial" };
        var saved = new Dictionary<string, string>();
        foreach (string name in names) saved[name] = Environment.GetEnvironmentVariable(name);
        try
        {
            string desktop = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
            if (String.IsNullOrWhiteSpace(desktop)) throw new Exception("Windows did not provide a DesktopDirectory for this test.");
            desktop = Path.GetFullPath(desktop).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
            string simulatedCloudRoot = Path.GetDirectoryName(desktop);
            if (String.IsNullOrWhiteSpace(simulatedCloudRoot)) throw new Exception("DesktopDirectory needs a parent directory for the simulated OneDrive policy tests.");
            foreach (string name in names) Environment.SetEnvironmentVariable(name, simulatedCloudRoot, EnvironmentVariableTarget.Process);

                Accept(delegate { ComparisonEngine.RequireOutputPath(desktop); }, "actual Desktop folder accepted");
                Accept(delegate { ComparisonEngine.RequireOutputPath(Path.Combine(desktop, "Results", "comparison.docx")); }, "Desktop descendants accepted");
                Accept(delegate { ComparisonEngine.RequireOutputPath(Path.Combine(desktop.ToUpperInvariant(), "comparison.docx")); }, "Desktop comparison is case insensitive");
                Accept(delegate { ComparisonEngine.RequireOutputPath(Path.Combine(desktop, "Results", "..", "comparison.docx")); }, "normalized path remaining inside Desktop accepted");
                Accept(delegate { ComparisonEngine.RequireOutputPath(Path.Combine(desktop + "-Other", "comparison.docx")); }, "other cloud output accepted");
                Accept(delegate { ComparisonEngine.RequireOutputPath(Path.Combine(desktop, "..", "Elsewhere", "comparison.docx")); }, "normalized cloud output accepted");
                Accept(delegate { ComparisonEngine.RequireOutputPath(Path.Combine(simulatedCloudRoot, "OtherCloudFolder", "comparison.docx")); }, "OneDrive output accepted");
                Accept(delegate { ComparisonEngine.RequireOutputPath(simulatedCloudRoot); }, "OneDrive root output accepted");
                Reject(delegate { ComparisonEngine.RequireLocalPath(Path.Combine(simulatedCloudRoot, "PrivateWork")); }, "Temporary working files", "temporary work stays private and local");
                Accept(delegate { ComparisonEngine.RequireSourcePath(Path.Combine(simulatedCloudRoot, "OtherCloudFolder", "source.docx")); }, "OneDrive inputs accepted");

            // The simulated root may be the whole user profile on non-OneDrive PCs.
            // Clear it for this independent ordinary-local-path check; no files are created.
            foreach (string name in names) Environment.SetEnvironmentVariable(name, null, EnvironmentVariableTarget.Process);
            string local = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "LocalDocumentCompare", "Results");
            Accept(delegate { ComparisonEngine.RequireOutputPath(local); }, "ordinary local output accepted");
            Accept(delegate { ComparisonEngine.RequireLocalPath(Path.Combine(local, "PrivateWork")); }, "ordinary local temporary work accepted");
            Accept(delegate { ComparisonEngine.RequireOutputPath(@"\\server\share\comparison.docx"); }, "UNC output accepted");
            Accept(delegate { ComparisonEngine.RequireSourcePath(@"\\server\share\source.docx"); }, "UNC source accepted");
            Accept(delegate { ComparisonEngine.RequireSourcePath(@"//server/share/source.docx"); }, "UNC forward slashes accepted");
            Reject(delegate { ComparisonEngine.RequireLocalPath(@"\\server\share\work"); }, "Temporary working files", "UNC private work rejected");
            foreach (string invalid in new[] { "relative.docx", @"C:relative.docx", @"\relative.docx", @"\\server", @"\\server\", "https://company.example/doc.docx", @"\\?\C:\file.docx", @"\\.\pipe\compare", "" })
            {
                string value = invalid;
                Reject(delegate { ComparisonEngine.RequireSourcePath(value); }, "file or folder path", "invalid source path rejected");
                Reject(delegate { ComparisonEngine.RequireOutputPath(value); }, "file or folder path", "invalid output path rejected");
            }

            int mappedDrives = 0;
            foreach (DriveInfo drive in DriveInfo.GetDrives())
            {
                if (drive.DriveType != DriveType.Network) continue;
                string mappedOutput = Path.Combine(drive.Name, "comparison.docx");
                Accept(delegate { ComparisonEngine.RequireOutputPath(mappedOutput); }, "mapped network output accepted");
                Accept(delegate { ComparisonEngine.RequireSourcePath(mappedOutput); }, "mapped network source accepted");
                Reject(delegate { ComparisonEngine.RequireLocalPath(mappedOutput); }, "Temporary working files", "mapped network private work rejected");
                mappedDrives++;
            }
            Console.WriteLine("PASS: " + checks + " Desktop/shared/output/input/private-work path checks; mapped network drives checked: " + mappedDrives + ". No files were created.");
            return 0;
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine("FAIL: " + ex.Message);
            return 1;
        }
        finally
        {
            foreach (string name in names) Environment.SetEnvironmentVariable(name, saved[name], EnvironmentVariableTarget.Process);
        }
    }

    private static void Accept(Action action, string name)
    {
        try { action(); }
        catch (Exception ex) { throw new Exception(name + ": unexpectedly rejected (" + ex.GetType().Name + ")."); }
        checks++;
    }

    private static void Reject(Action action, string expectedMessage, string name)
    {
        try { action(); }
        catch (Exception ex)
        {
            if (ex.Message.IndexOf(expectedMessage, StringComparison.OrdinalIgnoreCase) < 0)
                throw new Exception(name + ": rejected for an unexpected reason (" + ex.GetType().Name + ").");
            checks++;
            return;
        }
        throw new Exception(name + ": unexpectedly accepted.");
    }
}
