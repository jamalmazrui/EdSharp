using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Apress.VisualCSharpRecipes.Chapter16
{
    class Recipe16_13
    {
        static void Main(string[] args)
        {
            // define a numeric data source
            int[] ds1 = { 1, 23, 37, 49, 143 };

            // perform a query that shares a calculated value
            IEnumerable<double> result1 = from e in ds1
                                      let avg = ds1.Average()
                                      where (e < avg)
                                      select (e + avg);

            Console.WriteLine("Query using shared value");
            foreach (double element in result1)
            {
                Console.WriteLine("Result element {0}", element);
            }

            // Wait to continue.
            Console.WriteLine("\nMain method complete. Press Enter");
            Console.ReadLine();
        }
    }
}
