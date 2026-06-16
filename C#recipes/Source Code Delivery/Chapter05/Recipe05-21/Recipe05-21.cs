using System;
using System.IO;

namespace Apress.VisualCSharpRecipes.Chapter05
{
    static class Recipe05_21
    {
        static void Main()
        {
            string tempFile = Path.GetRandomFileName();

            Console.WriteLine("Using " + tempFile);

            using (FileStream fs = new FileStream(tempFile, FileMode.OpenOrCreate))
            {
                // (Write some data.)
            }

            // Now delete the file.
            File.Delete(tempFile);

            // Wait to continue.
            Console.WriteLine(Environment.NewLine);
            Console.WriteLine("Main method complete. Press Enter.");
            Console.ReadLine();
        }
    }
}