using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Apress.VisualCSharpRecipes.Chapter16
{
    static class LINQExtensions
    {
        public static IEnumerable<string> RemoveFirstAndLast(this IEnumerable<string> source)
        {
            return source.Skip(1).Take(source.Count() - 2);
        }
    }

    class Recipe16_14
    {
        static void Main(string[] args)
        {
            // create the data sources
            string[] ds1 = {"apple", "banana", "pear", "fig"};
            IList<string> ds2 = new List<string> { "apple", "banana", "pear", "fig" };

            Console.WriteLine("Extension method used on string[]");
            IEnumerable<string> result1 = ds1.RemoveFirstAndLast();
            foreach (string element in result1)
            {
                Console.WriteLine("Result: {0}", element);
            }

            Console.WriteLine("\nExtension method used on IList<string>");
            IEnumerable<string> result2 = ds1.RemoveFirstAndLast();
            foreach (string element in result2)
            {
                Console.WriteLine("Result: {0}", element);
            }
               
            // Wait to continue.
            Console.WriteLine("\nMain method complete. Press Enter");
            Console.ReadLine();
        }
    }
}
