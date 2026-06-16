using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Dynamic;

namespace Apress.VisualCSharpRecipes.Chapter13
{
    class myType
    {
        public myType(string strval)
        {
            str = strval;
        }

        public string str {get; set;}

        public int countVowels()
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

    class Recipe13_16
    {
        static void Main(string[] args)
        {
            // create a dynamic type
            dynamic dynInstance = new myType("The quick brown fox jumped over the...");
            // call the countVowels method
            int vowels = dynInstance.countVowels();
            Console.WriteLine("There are {0} vowels", vowels);
            // call a method that does not exist
            dynInstance.thisMethodDoesNotExist();
        }
    }
}
