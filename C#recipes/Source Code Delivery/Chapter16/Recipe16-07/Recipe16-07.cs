using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Apress.VisualCSharpRecipes.Chapter16
{
    class Recipe16_07
    {
        static void Main(string[] args)
        {
            // create the data sources
            string[] datasource1 = { "apple", "orange", "cherry", "pear" };
            int[]    datasource2 = { 21, 42, 37 };

            // perform the linq query
            var result = from e in datasource1
                         from f in datasource2
                         select new
                         {
                             e,
                             f
                         };

            // print the results
            foreach (var element in result)
            {
                Console.WriteLine("{0}, {1}", element.e, element.f);
            }

            // Wait to continue.
            Console.WriteLine("\nMain method complete. Press Enter");
            Console.ReadLine();
        } 
    }
}
