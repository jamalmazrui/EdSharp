using System;
using System.IO;

namespace Apress.VisualCSharpRecipes.Chapter02
{
    class Recipe02_03
    {
        // Create a byte array from a decimal.
        public static byte[] DecimalToByteArray (decimal src) 
        {
            // Create a MemoryStream as a buffer to hold the binary data.
            using (MemoryStream stream = new MemoryStream()) 
            {
                // Create a BinaryWriter to write binary data to the stream.
                using (BinaryWriter writer = new BinaryWriter(stream)) 
                {
                    // Write the decimal to the BinaryWriter/MemoryStream.
                    writer.Write(src);

                    // Return the byte representation of the decimal.
                    return stream.ToArray();
                }
            }
        }

        // Create a decimal from a byte array.
        public static decimal ByteArrayToDecimal (byte[] src) 
        {
            // Create a MemoryStream containing the byte array.
            using (MemoryStream stream = new MemoryStream(src)) 
            {
                // Create a BinaryReader to read the decimal from the stream.
                using (BinaryReader reader = new BinaryReader(stream)) 
                {
                    // Read and return the decimal from the 
                    // BinaryReader/MemoryStream.
                    return reader.ReadDecimal();
                }
            }
        }

        public static void Main() 
        {
            byte[] b = null;

            // Convert a bool to a byte array and display.
            b = BitConverter.GetBytes(true);
            Console.WriteLine(BitConverter.ToString(b));

            // Convert a byte array to a bool and display.
            Console.WriteLine(BitConverter.ToBoolean(b,0));

            // Convert an int to a byte array and display.
            b = BitConverter.GetBytes(3678);
            Console.WriteLine(BitConverter.ToString(b));

            // Convert a byte array to an int and display.
            Console.WriteLine(BitConverter.ToInt32(b,0));

            // Convert a decimal to a byte array and display.
            b = DecimalToByteArray(285998345545.563846696m);
            Console.WriteLine(BitConverter.ToString(b));

            // Convert a byte array to a decimal and display.
            Console.WriteLine(ByteArrayToDecimal(b));

            // Wait to continue.
            Console.WriteLine("Main method complete. Press Enter");
            Console.ReadLine();
        }
    }
}