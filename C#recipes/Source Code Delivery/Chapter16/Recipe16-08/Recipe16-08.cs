using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Apress.VisualCSharpRecipes.Chapter16
{
    class Recipe16_08
    {
        static void Main(string[] args)
        {
            // create the data sources
            string[] datasource1 = { "apple", "orange", "cherry", "pear" };
            string[] datasource2 = { "banana", "kiwi", "fig" };

            // perform the linq query
            IEnumerable<string> result = from e in datasource1.Concat<string>(datasource2)
                select e;

            // print the results
            foreach (string element in result)
            {
                Console.WriteLine(element);
            }

            // Wait to continue.
            Console.WriteLine("\nMain method complete. Press Enter");
            Console.ReadLine();
        }
    }
}
