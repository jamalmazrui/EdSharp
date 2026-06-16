using System;
using System.Text;
using System.Security.Cryptography;

namespace Apress.VisualCSharpRecipes.Chapter11
{
    class Recipe11_14
    {
        public static void Main(string[] args)
        {
            // Create a HashAlgorithm of the type specified by the first
            // command-line argument.
            HashAlgorithm hashAlg = null;
            if (args[0].CompareTo("SHAManaged1") == 0)
            {
                hashAlg = new SHA1Managed();
            }
            else
            {
                hashAlg = HashAlgorithm.Create(args[0]);
            }

            using (hashAlg)
            {
                // Convert the password string, provided as the second 
                // command-line argument, to an array of bytes.
                byte[] pwordData = Encoding.Default.GetBytes(args[1]);

                // Generate the hash code of the password.
                byte[] hash = hashAlg.ComputeHash(pwordData);

                // Display the hash code of the password to the console.
                Console.WriteLine(BitConverter.ToString(hash));

                // Wait to continue.
                Console.WriteLine("\nMain method complete. Press Enter.");
                Console.ReadLine();
            }
        }
    }
}
