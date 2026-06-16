using System;
using System.IO;

namespace Apress.VisualCSharpRecipes.Chapter05
{
    static class Recipe05_02
    {
        static void Main()
        {
            // This file has the archive and read-only attributes.
            FileInfo file = new FileInfo(@"C:\Windows\win.ini");

            // This displays the attributes
            Console.WriteLine(file.Attributes.ToString());

            // This test fails, because other attributes are set.
            if (file.Attributes == FileAttributes.ReadOnly)
            {
                Console.WriteLine("File is read-only (faulty test).");
            }

            // This test succeeds, because it filters out just the
            // read-only attribute.
            if ((file.Attributes & FileAttributes.ReadOnly) ==
              FileAttributes.ReadOnly)
            {
                Console.WriteLine("File is read-only (correct test).");
            }

            // Wait to continue.
            Console.WriteLine(Environment.NewLine);
            Console.WriteLine("Main method complete. Press Enter.");
            Console.ReadLine();
        }
    }
}
