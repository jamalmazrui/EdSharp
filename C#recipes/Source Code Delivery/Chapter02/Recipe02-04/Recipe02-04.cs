using System;
using System.IO;
using System.Text;

namespace Apress.VisualCSharpRecipes.Chapter02
{
    class Recipe02_04
    {
        // Create a byte array from a decimal.
        public static byte[] DecimalToByteArray (decimal src) 
        {
            // Create a MemoryStream as a buffer to hold the binary data.
            using (MemoryStream stream = new MemoryStream()) 
            {
                // Create a BinaryWriter to write binary data the stream.
                using (BinaryWriter writer = new BinaryWriter(stream)) 
                {
                    // Write the decimal to the BinaryWriter/MemoryStream.
                    writer.Write(src);

                    // Return the byte representation of the decimal.
                    return stream.ToArray();
                }
            }
        }

        // Create a decimal from a byte array
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

        // Base64 encode a Unicode string as a string.
        public static string StringToBase64 (string src) 
        {
            // Get a byte representation of the source string.
            byte[] b = Encoding.Unicode.GetBytes(src);

            // Return the Base64-encoded string.
            return Convert.ToBase64String(b);
        }

        // Decode a Base64-encoded Unicode string from a string.
        public static string Base64ToString (string src) 
        {
            // Decode the Base64-encoded string to a byte array.
            byte[] b = Convert.FromBase64String(src);

            // Return the decoded Unicode string.
            return Encoding.Unicode.GetString(b);
        }

        // Base64 encode a decimal.
        public static string DecimalToBase64 (decimal src) 
        {
            // Get a byte representation of the decimal.
            byte[] b = DecimalToByteArray(src);

            // Return the Base64-encoded decimal.
            return Convert.ToBase64String(b);
        }

        // Decode a Base64-encoded decimal.
        public static decimal Base64ToDecimal (string src) 
        {
            // Decode the Base64-encoded decimal to a byte array.
            byte[] b = Convert.FromBase64String(src);

            // Return the decoded decimal.
            return ByteArrayToDecimal(b);
        }

        // Base64 encode an int.
        public static string IntToBase64 (int src) 
        {
            // Get a byte representation of the int.
            byte[] b = BitConverter.GetBytes(src);

            // Return the Base64-encoded int.
            return Convert.ToBase64String(b);
        }

        // Decode a Base64-encoded int.
        public static int Base64ToInt (string src) 
        {
            // Decode the Base64-encoded int to a byte array.
            byte[] b = Convert.FromBase64String(src);

            // Return the decoded int.
            return BitConverter.ToInt32(b,0);
        }

        public static void Main() 
        {
            // Encode and decode a general byte array. Need to create a char[]
            // to hold the Base64 encoded data. The size of the char[] must 
            // be at least 4/3 the size of the source byte[] and must be 
            // divisible by 4.
            byte[] data = { 0x04, 0x43, 0x5F, 0xFF, 0x0, 0xF0, 0x2D, 0x62, 0x78,  
                0x22, 0x15, 0x51, 0x5A, 0xD6, 0x0C, 0x59, 0x36, 0x63, 0xBD, 0xC2, 
                0xD5, 0x0F, 0x8C, 0xF5, 0xCA, 0x0C};

            char[] base64data = 
                new char[(int)(Math.Ceiling((double)data.Length / 3) * 4)];

            Console.WriteLine("\nByte array encoding/decoding");
            Convert.ToBase64CharArray(data, 0, data.Length, base64data, 0);
            Console.WriteLine(new String(base64data));
            Console.WriteLine(BitConverter.ToString(
                Convert.FromBase64CharArray(base64data, 0, base64data.Length)));

            // Encode and decode a string.
            Console.WriteLine("\nString encoding/decoding");
            Console.WriteLine(StringToBase64
                ("Welcome to Visual C# Recipes from Apress"));
            Console.WriteLine(Base64ToString("VwBlAGwAYwBvAG0AZQAgAHQAbwA" +
                "gAFYAaQBzAHUAYQBsACAAQwAjACAAUgBlAGMAaQBwAGUAcwAgAGYAcgB" +
                "vAG0AIABBAHAAcgBlAHMAcwA="));

            // Encode and decode a decimal.
            Console.WriteLine("\nDecimal encoding/decoding");
            Console.WriteLine(DecimalToBase64(285998345545.563846696m));
            Console.WriteLine(Base64ToDecimal("KDjBUP07BoEPAAAAAAAJAA=="));

            // Encode and decode an int.
            Console.WriteLine("\nInteger encoding/decoding");
            Console.WriteLine(IntToBase64(35789));
            Console.WriteLine(Base64ToInt("zYsAAA=="));

            // Wait to continue.
            Console.WriteLine("\nMain method complete. Press Enter");
            Console.ReadLine();
        }
    }
}