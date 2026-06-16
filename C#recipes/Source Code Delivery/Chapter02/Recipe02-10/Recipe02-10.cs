using System;
using System.Collections;

namespace Apress.VisualCSharpRecipes.Chapter02
{
    class Recipe02_10
    {
        public static void Main()
        {
            // Create a new ArrayList and populate it.
            ArrayList list = new ArrayList(5);
            list.Add("Brenda");
            list.Add("George");
            list.Add("Justin");
            list.Add("Shaun");
            list.Add("Meaghan");

            // Create a string array and use the ICollection.CopyTo method 
            // to copy the contents of the ArrayList.
            string[] array1 = new string[list.Count];
            list.CopyTo(array1, 0);

            // Use ArrayList.ToArray to create an object array from the 
            // contents of the collection.
            object[] array2 = list.ToArray();

            // Use ArrayList.ToArray to create a strongly typed string
            // array from the contents of the collection.
            string[] array3 = (string[])list.ToArray(typeof(String));

            // Display the contents of the 3 arrays.
            Console.WriteLine("Array 1:");
            foreach (string s in array1) 
            { 
                Console.WriteLine("\t{0}",s); 
            }

            Console.WriteLine("Array 2:");
            foreach (string s in array2) 
            {
                Console.WriteLine("\t{0}", s);
            }

            Console.WriteLine("Array 3:");
            foreach (string s in array3) 
            {
                Console.WriteLine("\t{0}", s);
            }

            // Wait to continue.
            Console.WriteLine("\nMain method complete. Press Enter");
            Console.ReadLine();
        }
    }
}
