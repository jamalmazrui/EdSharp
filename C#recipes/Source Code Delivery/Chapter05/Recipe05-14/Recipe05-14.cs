using System;
using System.IO;

namespace Apress.VisualCSharpRecipes.Chapter05
{
    static class Recipe05_14
    {
        static void Main()
        {
            Console.WriteLine("Using: " + Directory.GetCurrentDirectory());
            Console.WriteLine("The relative path 'file.txt' " +
              "will automatically become: '" +
              Path.GetFullPath("file.txt") + "'");

            Console.WriteLine();

            Console.WriteLine("Changing current directory to c:\\");
            Directory.SetCurrentDirectory(@"c:\");

            Console.WriteLine("Now the relative path 'file.txt' " +
              "will automatically become '" +
              Path.GetFullPath("file.txt") + "'");
            
            // Wait to continue.
            Console.WriteLine(Environment.NewLine);
            Console.WriteLine("Main method complete. Press Enter.");
            Console.ReadLine();
        }
    }
}
