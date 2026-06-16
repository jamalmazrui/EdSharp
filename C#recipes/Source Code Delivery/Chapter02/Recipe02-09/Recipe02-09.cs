using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

namespace Apress.VisualCSharpRecipes.Chapter02
{
    class Recipe02_09
    {
        public static void Main()
        {
            // Create a new array and populate it.
            int[] array = { 4, 2, 9, 3 };

            // created a new, sorted array
            array = Enumerable.OrderBy(array, e => e).ToArray<int>();

            // Display the contents of the sorted array.
            foreach (int i in array) { 
                Console.WriteLine(i); 
            }

            // create a list and populate it
            List<string> list = new List<string>();
            list.Add("Michael");
            list.Add("Kate");
            list.Add("Andrea");
            list.Add("Angus");

            // enumerate the sorted contents of the list
            Console.WriteLine("\nList sorted by content");
            foreach (string person in Enumerable.OrderBy(list, e => e))
            {
                Console.WriteLine(person);
            }

            // sort and enumerate based on a property
            Console.WriteLine("\nList sorted by length property");
            foreach (string person in Enumerable.OrderBy(list, e => e.Length))
            {
                Console.WriteLine(person);
            }

            // Create a new ArrayList and populate it.
            ArrayList arraylist = new ArrayList(4);
            arraylist.Add("Michael");
            arraylist.Add("Kate");
            arraylist.Add("Andrea");
            arraylist.Add("Angus");

            // Sort the ArrayList.
            arraylist.Sort();

            // Display the contents of the sorted ArrayList.
            Console.WriteLine("\nArraylist sorted by content");
            foreach (string s in list) { 
                Console.WriteLine(s); 
            }

            // Wait to continue.
            Console.WriteLine("\nMain method complete. Press Enter");
            Console.ReadLine();
        }
    }
}
