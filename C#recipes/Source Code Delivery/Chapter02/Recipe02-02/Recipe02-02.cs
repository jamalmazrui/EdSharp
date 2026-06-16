using System;
using System.IO;
using System.Text;

namespace Apress.VisualCSharpRecipes.Chapter02
{
    class Recipe02_02
    {
        public static void Main() 
        {        
            // Create a file to hold the output.
            using (StreamWriter output = new StreamWriter("output.txt")) 
            {
                // Create and write a string containing the symbol for Pi.
                string srcString = "Area = \u03A0r^2";
                output.WriteLine("Source Text : " + srcString);

                // Write the UTF-16 encoded bytes of the source string.
                byte[] utf16String = Encoding.Unicode.GetBytes(srcString);
                output.WriteLine("UTF-16 Bytes: {0}", 
                    BitConverter.ToString(utf16String));

                // Convert the UTF-16 encoded source string to UTF-8 and ASCII.
                byte[] utf8String = Encoding.UTF8.GetBytes(srcString);
                byte[] asciiString = Encoding.ASCII.GetBytes(srcString);
                
                // Write the UTF-8 and ASCII encoded byte arrays. 
                output.WriteLine("UTF-8  Bytes: {0}", 
                    BitConverter.ToString(utf8String));
                output.WriteLine("ASCII  Bytes: {0}", 
                    BitConverter.ToString(asciiString));

                // Convert UTF-8 and ASCII encoded bytes back to UTF-16 encoded  
                // string and write.
                output.WriteLine("UTF-8  Text : {0}", 
                    Encoding.UTF8.GetString(utf8String));
                output.WriteLine("ASCII  Text : {0}", 
                    Encoding.ASCII.GetString(asciiString));
            }

            // Wait to continue.
            Console.WriteLine("Main method complete. Press Enter");
            Console.ReadLine();
        }
    }
}

