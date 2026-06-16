using System;
using System.IO;
using System.Text;
using System.Security.Cryptography;

namespace Apress.VisualCSharpRecipes.Chapter11
{
    class Recipe11_17
    {
        public static void Main(string[] args)
        {
            // Create a byte array from the key string, which is the 
            // second command-line argument.
            byte[] key = Encoding.Unicode.GetBytes(args[2]);

            // Create a KeyedHashAlgorithm derived object to generate the keyed 
            // hash code for the input file. Pass the byte array representing the  
            // key to the constructor.
            using (KeyedHashAlgorithm hashAlg = KeyedHashAlgorithm.Create(args[1]))
            {
                // Assign the key.
                hashAlg.Key = key;

                // Open a FileStream to read the input file; the file name is 
                // specified by the first command-line argument.
                using (Stream file = 
                    new FileStream(args[0], FileMode.Open,FileAccess.Read))
                {
                    // Generate the keyed hash code of the file’s contents.
                    byte[] hash = hashAlg.ComputeHash(file);

                    // Display the keyed hash code to the console.
                    Console.WriteLine(BitConverter.ToString(hash));
                }
            }

            // Wait to continue.
            Console.WriteLine("\nMain method complete. Press Enter.");
            Console.ReadLine();
        }
    }
}