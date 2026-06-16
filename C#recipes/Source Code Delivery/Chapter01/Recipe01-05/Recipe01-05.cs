using System;

namespace Apress.VisualCSharpRecipes.Chapter01
{
    class Recipe01_05
    {
        public static void Main(string[] args)
        {
            // Step through the command-line arguments.
            foreach (string s in args)
            {
                Console.WriteLine(s);
            }

            // Alternatively, access the command-line arguments directly.
            Console.WriteLine(Environment.CommandLine);

            foreach (string s in Environment.GetCommandLineArgs())
            {
                Console.WriteLine(s);
            }

            // Wait to continue.
            Console.WriteLine("\nMain method complete. Press Enter.");
            Console.ReadLine();
        }
    }
}
