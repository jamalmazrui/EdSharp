using System;
using System.Text;

namespace Apress.VisualCSharpRecipes.Chapter02
{
    class Recipe02_01
    {
        public static string ReverseString(string str) 
        {
            // Make sure we have a reversible string.
            if (str == null || str.Length <= 1)
            {
                return str;
            }

            // Create a StringBuilder object with the required capacity.
            StringBuilder revStr = new StringBuilder(str.Length);

            // Loop backward through the source string one character at a time and append 
            // each character to the StringBuilder.
            for (int count = str.Length - 1; count > -1; count--)
            {
                revStr.Append(str[count]);
            }

            // Return the reversed string.
            return revStr.ToString();
        }

        public static void Main()
        {
            Console.WriteLine(ReverseString("Madam Im Adam"));

            Console.WriteLine(ReverseString("The quick brown fox jumped over the lazy dog."));

            // Wait to continue.
            Console.WriteLine("\nMain method complete. Press Enter");
            Console.ReadLine();
        }
    }
}
