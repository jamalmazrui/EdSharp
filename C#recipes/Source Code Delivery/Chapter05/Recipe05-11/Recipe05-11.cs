using System;
using System.IO;
using System.Security.Cryptography;

namespace Apress.VisualCSharpRecipes.Chapter05
{
    static class Recipe05_11
    {
        static void Main(string[] args)
        {
            if (args.Length != 2)
            {
                Console.WriteLine("USAGE:  Recipe05_11 [fileName] [fileName]");
                return;
            }

            Console.WriteLine("Comparing " + args[0] + " and " + args[1]);

            // Create the hashing object.
            using (HashAlgorithm hashAlg = HashAlgorithm.Create())
            {
                using (FileStream fsA = new FileStream(args[0], FileMode.Open),
                    fsB = new FileStream(args[1], FileMode.Open))
                {
                    // Calculate the hash for the files.
                    byte[] hashBytesA = hashAlg.ComputeHash(fsA);
                    byte[] hashBytesB = hashAlg.ComputeHash(fsB);

                    // Compare the hashes.
                    if (BitConverter.ToString(hashBytesA) ==
                        BitConverter.ToString(hashBytesB))
                    {
                        Console.WriteLine("Files match.");
                    }
                    else
                    {
                        Console.WriteLine("No match.");
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
