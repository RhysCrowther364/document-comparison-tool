using System;
using System.Runtime.InteropServices;
public static class WordProbe
{
    [STAThread] public static int Main()
    {
        dynamic word = null;
        dynamic document = null;
        try
        {
            Console.WriteLine("Create");
            word = Activator.CreateInstance(Type.GetTypeFromProgID("Word.Application"));
            Console.WriteLine("Version: " + word.Version);
            Console.WriteLine("Documents");
            dynamic documents = word.Documents;
            Console.WriteLine("Add");
            document = documents.Add();
            Console.WriteLine("Windows");
            dynamic windows = document.Windows;
            Console.WriteLine("Item");
            dynamic window = windows.Item(1);
            Console.WriteLine("Hwnd: " + window.Hwnd);
            Console.WriteLine("Close");
            document.Close(SaveChanges: 0);
            Console.WriteLine("Quit");
            word.Quit(SaveChanges: 0);
            Console.WriteLine("Done");
            return 0;
        }
        catch (Exception ex) { Console.WriteLine(ex.ToString()); return 1; }
    }
}
