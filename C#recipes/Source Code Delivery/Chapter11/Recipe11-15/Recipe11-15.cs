using System;
using System.IO;
using System.Security.Cryptography;

namespace Apress.VisualCSharpRecipes.Chapter11
{
    class Recipe11_15
    {
        public static void Main(string[] args)
        {
            // Create a HashAlgorithm of the type specified by the first
            // command-line argument.
            using (HashAlgorithm hashAlg = HashAlgorithm.Create(args[0]))
            {
                // Open a FileStream to the file specified by the second
                // command-line argument.
                using (Stream file = 
                    new FileStream(args[1], FileMode.Open, FileAccess.Read))
                {
                    // Generate the hash code of the file’s contents.
                    byte[] hash = hashAlg.ComputeHash(file);

                    // Display the hash code of the file to the console.
                    Console.WriteLine(BitConverter.ToString(hash));
                }

                // Wait to continue.
                Console.WriteLine("\nMain method complete. Press Enter.");
                Console.ReadLine();
            }
        }
    }
}
