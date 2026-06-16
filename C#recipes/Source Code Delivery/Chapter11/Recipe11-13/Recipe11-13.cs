using System;
using System.Security.Cryptography;

namespace Apress.VisualCSharpRecipes.Chapter11
{
    class Recipe11_13 
    {
        public static void Main() {
        
            // Create a byte array to hold the random data.
            byte[] number = new byte[32];
            
            // Instantiate the default random number generator.
            RandomNumberGenerator rng = RandomNumberGenerator.Create();
            
            // Generate 32 bytes of random data.
            rng.GetBytes(number);
            
            // Display the random number.
            Console.WriteLine(BitConverter.ToString(number));

            // Wait to continue.
            Console.WriteLine("\nMain method complete. Press Enter.");
            Console.ReadLine();
        }
    }
}
