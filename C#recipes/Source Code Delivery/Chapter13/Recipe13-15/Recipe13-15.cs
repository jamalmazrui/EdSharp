using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Apress.VisualCSharpRecipes.Chapter13.Extensions;

namespace Apress.VisualCSharpRecipes.Chapter13.Extensions
{
    public static class MyStringExtentions
    {
        public static string toMixedCase(this String str)
        {
            StringBuilder builder = new StringBuilder(str.Length);
            for (int i = 0; i < str.Length; i += 2)
            {
                builder.Append(str.ToLower()[i]);
                builder.Append(str.ToUpper()[i + 1]);
            }
            return builder.ToString();
        }

        public static int countVowels(this String str)
        {
            char[] vowels = { 'a', 'e', 'i', 'o', 'u' };
            int vowelcount = 0;
            foreach (char c in str)
            {
                if (vowels.Contains(c))
                {
                    vowelcount++;
                }
            }
            return vowelcount;
        }
    }
}

namespace Apress.VisualCSharpRecipes.Chapter13
{
    class Recipe13_15
    {
        static void Main(string[] args)
        {
            string str = "The quick brown fox jumped over the...";
            Console.WriteLine(str.toMixedCase());
            Console.WriteLine("There are {0} vowels", str.countVowels());

            // Wait to continue.
            Console.WriteLine(Environment.NewLine);
            Console.WriteLine("Main method complete. Press Enter");
            Console.ReadLine();
        }
    }
}
