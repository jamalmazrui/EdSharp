using System;
using System.IO;
using System.Text;

namespace Apress.VisualCSharpRecipes.Chapter05
{
    static class Recipe05_07
    {
        static void Main()
        {
            // Create a new file.
            using (FileStream fs = new FileStream("test.txt", FileMode.Create))
            {
                // Create a writer and specify the encoding.
                // The default (UTF-8) supports special Unicode characters,
                // but encodes all standard characters in the same way as
                // ASCII encoding.
                using (StreamWriter w = new StreamWriter(fs, Encoding.UTF8))
                {
                    // Write a decimal, string, and char.
                    w.WriteLine(124.23M);
                    w.WriteLine("Test string");
                    w.WriteLine('!');
                }
            }
            Console.WriteLine("Press Enter to read the information.");
            Console.ReadLine();

            // Open the file in read-only mode.
            using (FileStream fs = new FileStream("test.txt", FileMode.Open))
            {
                using (StreamReader r = new StreamReader(fs, Encoding.UTF8))
                {
                    // Read the data and convert it to the appropriate data type.
                    Console.WriteLine(Decimal.Parse(r.ReadLine()));
                    Console.WriteLine(r.ReadLine());
                    Console.WriteLine(Char.Parse(r.ReadLine()));
                }
            }

            // Wait to continue.
            Console.WriteLine(Environment.NewLine);
            Console.WriteLine("Main method complete. Press Enter.");
            Console.ReadLine();
        }
    }
}
