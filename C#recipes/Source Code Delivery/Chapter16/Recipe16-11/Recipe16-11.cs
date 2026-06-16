using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Apress.VisualCSharpRecipes.Chapter16
{
    class Recipe16_11
    {
        static void Main(string[] args)
        {
            // create the first data source
            string[] ds1 = { "apple", "cherry", "pear" };
            // create a data source with the same elements
            // in the same order
            string[] ds2 = { "apple", "cherry", "pear" };
            // create a data source with the 
            // same elements in a different order
            string[] ds3 = { "pear", "cherry", "apple" };
            // create a data source with different elements
            string[] ds4 = { "apricot", "cranberry", "plum" };
            
            // perform the comparisons
            Console.WriteLine("Using standard comparer");
            Console.WriteLine("DS1 == DS2? {0}", ds1.SequenceEqual(ds2));
            Console.WriteLine("DS1 == DS3? {0}", ds1.SequenceEqual(ds3));
            Console.WriteLine("DS1 == DS4? {0}", ds1.SequenceEqual(ds4));

            // create the custom comparer
            MyComparer comparer = new MyComparer();

            Console.WriteLine("\nUsing custom comparer");
            Console.WriteLine("DS1 == DS2? {0}", ds1.SequenceEqual(ds2, comparer));
            Console.WriteLine("DS1 == DS3? {0}", ds1.SequenceEqual(ds3, comparer));
            Console.WriteLine("DS1 == DS4? {0}", ds1.SequenceEqual(ds4, comparer));

            // Wait to continue.
            Console.WriteLine("\nMain method complete. Press Enter");
            Console.ReadLine();
        }
    }
    class MyComparer : IEqualityComparer<string>
    {
        public bool Equals(string first, string second)
        {
            return first[0] == second[0];
        }

        public int GetHashCode(string str)
        {
            return str[0].GetHashCode();
        }
    }
}
