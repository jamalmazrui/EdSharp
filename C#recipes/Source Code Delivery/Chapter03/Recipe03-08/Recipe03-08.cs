using System;
using System.Reflection;
using System.Collections;

namespace Apress.VisualCSharpRecipes.Chapter03
{
    public class ListModifier
    {
        public ListModifier() 
        {
            // Get the list from the data cache.
            ArrayList list = (ArrayList)AppDomain.CurrentDomain.GetData("Pets");

            // Modify the list. 
            list.Add("turtle");
        }
    }

    class Recipe03_08
    {
        public static void Main() 
        {
            // Create a new application domain.
            AppDomain domain = AppDomain.CreateDomain("Test");

            // Create an ArrayList and populate with information.
            ArrayList list = new ArrayList();
            list.Add("dog");
            list.Add("cat");
            list.Add("fish");

            // Place the list in the data cache of the new application domain.
            domain.SetData("Pets", list);

            // Instantiate a ListModifier in the new application domain.
            domain.CreateInstance("Recipe03-08", 
                "Apress.VisualCSharpRecipes.Chapter03.ListModifier");

            // Display the contents of the original list.
            Console.WriteLine("Original list contents:");
            foreach (string s in list)
            {
                Console.WriteLine("  - " + s);
            }

            // Get the list and display its contents.
            Console.WriteLine("\nModified list contents:");
            foreach (string s in (ArrayList)domain.GetData("Pets"))
            {
                Console.WriteLine("  - " + s);
            }

            // Wait to continue.
            Console.WriteLine("\nMain method complete. Press Enter.");
            Console.ReadLine();
        }
    }
}