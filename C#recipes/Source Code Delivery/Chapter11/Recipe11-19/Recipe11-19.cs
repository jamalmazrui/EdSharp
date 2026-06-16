using System;
using System.Text;
using System.Security.Cryptography;

namespace Apress.VisualCSharpRecipes.Chapter11
{
    class Recipe11_19
    {
        public static void Main()
        {
            // Read the string from the console.
            Console.Write("Enter the string to encrypt: ");
            string str = Console.ReadLine();

            // Create a byte array of entropy to use in the encryption process.
            byte[] entropy = { 0, 1, 2, 3, 4, 5, 6, 7, 8 };

            // Encrypt the entered string after converting it to 
            // a byte array. Use CurrentUser scope so that only
            // the current user can decrypt the data.
            byte[] enc = ProtectedData.Protect(Encoding.Unicode.GetBytes(str),
                entropy, DataProtectionScope.LocalMachine);

            // Display the encrypted data to the console.
            Console.WriteLine("\nEncrypted string = {0}", 
                BitConverter.ToString(enc));

            // Attempt to decrypt the data using CurrentUser scope.
            byte[] dec = ProtectedData.Unprotect(enc,
                entropy, DataProtectionScope.CurrentUser);

            // Display the data decrypted using CurrentUser scope.
            Console.WriteLine("\nDecrypted data using CurrentUser scope = {0}",
                Encoding.Unicode.GetString(dec));

            // Wait to continue.
            Console.WriteLine("\nMain method complete. Press Enter.");
            Console.ReadLine();
        }
    }
}
