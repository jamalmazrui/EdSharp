using System;
using System.IO;

namespace Apress.VisualCSharpRecipes.Chapter05
{
    static class Recipe05_12
    {
        static void Main(string[] args)
        {
            string filename;
            filename = @"..\System\MyFile.txt";
            filename = Path.GetFileName(filename);

            Console.WriteLine(filename);

            filename = @"..\..\myfile.txt";
            string fullPath = @"c:\Temp";

            filename = Path.GetFileName(filename);
            fullPath = Path.Combine(fullPath, filename);

            Console.WriteLine(fullPath);
            Console.WriteLine(filename);

            // (fullPath is now "c:\Temp\myfile.txt")

            // Wait to continue.
            Console.WriteLine(Environment.NewLine);
            Console.WriteLine("Main method complete. Press Enter.");
            Console.ReadLine();
        }
    }
}