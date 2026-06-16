using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Apress.VisualCSharpRecipes.Chapter16
{
    class Recipe16_12
    {
        static void Main(string[] args)
        {
            // define a numeric data source
            int[] ds1 = { 1, 23, 37, 49, 143 };

            // use the standard aggregation methods
            Console.WriteLine("Standard aggregation methods");
            Console.WriteLine("Average: {0}", ds1.Average());
            Console.WriteLine("Count: {0}", ds1.Count());
            Console.WriteLine("Max: {0}", ds1.Max());
            Console.WriteLine("Min: {0}", ds1.Min());
            Console.WriteLine("Sum: {0}", ds1.Sum());

            // perform our own sum aggregation
            Console.WriteLine("\nCustom aggregation");
            Console.WriteLine(ds1.Aggregate((total, elem) => total += elem));

            // define a string data source
            string[] ds2 = { "apple", "pear", "cherry" };

            // perform a concat aggregation
            Console.WriteLine("\nString concatenation aggregation");
            Console.WriteLine(ds2.Aggregate((len, elem) => len += elem));

            // Wait to continue.
            Console.WriteLine("\nMain method complete. Press Enter");
            Console.ReadLine();
        }
    }
}
