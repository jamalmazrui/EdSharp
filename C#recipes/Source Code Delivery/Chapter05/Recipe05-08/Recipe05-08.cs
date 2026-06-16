using System;
using System.IO;

namespace Apress.VisualCSharpRecipes.Chapter05
{
    static class Recipe05_08
    {
        static void Main()
        {
            // Create a new file and writer.
            using (FileStream fs = new FileStream("test.bin", FileMode.Create))
            {
                using (BinaryWriter w = new BinaryWriter(fs))
                {
                    // Write a decimal, two strings, and a char.
                    w.Write(124.23M);
                    w.Write("Test string");
                    w.Write("Test string 2");
                    w.Write('!');
                }
            }
            Console.WriteLine("Press Enter to read the information.");
            Console.ReadLine();

            // Open the file in read-only mode.
            using (FileStream fs = new FileStream("test.bin", FileMode.Open))
            {
                // Display the raw information in the file.
                using (StreamReader sr = new StreamReader(fs))
                {
                    Console.WriteLine(sr.ReadToEnd());
                    Console.WriteLine();

                    // Read the data and convert it to the appropriate data type.
                    fs.Position = 0;
                    using (BinaryReader br = new BinaryReader(fs))
                    {
                        Console.WriteLine(br.ReadDecimal());
                        Console.WriteLine(br.ReadString());
                        Console.WriteLine(br.ReadString());
                        Console.WriteLine(br.ReadChar());
                    }
                }
            }

            // Wait to continue.
            Console.WriteLine(Environment.NewLine);
            Console.WriteLine("Main method complete. Press Enter.");
            Console.ReadLine();
        }
    }
}
