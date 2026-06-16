using System;
using System.IO;

namespace Apress.VisualCSharpRecipes.Chapter05
{
    static class Recipe05_13
    {
        static void Main(string[] args)
        {
            foreach (string arg in args)
            {
                Console.Write(arg);

                if (Directory.Exists(arg))
                {
                    Console.WriteLine(" is a directory");
                }
                else if (File.Exists(arg))
                {
                    Console.WriteLine(" is a file");
                }
                else
                {
                    Console.WriteLine(" does not exist");
                }
            }
            
            // Wait to continue.
            Console.WriteLine(Environment.NewLine);
            Console.WriteLine("Main method complete. Press Enter.");
            Console.ReadLine();
        }
    }
}
