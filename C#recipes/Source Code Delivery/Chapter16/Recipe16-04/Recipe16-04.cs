using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Apress.VisualCSharpRecipes.Chapter16
{
    class Recipe16_04
    {
        static void Main(string[] args)
        {
            string[] array = { "one", "two", "three", "four", "five" };

            IEnumerable<string> skipresult = from e in array.Skip<string>(2) select e;
            foreach (string str in skipresult)
            {
                Console.WriteLine("Result from skip filter: {0}", str);
            }

            IEnumerable<string> takeresult = from e in array.Take<string>(2) select e;
            foreach (string str in takeresult)
            {
                Console.WriteLine("Result from take filter: {0}", str);
            }

            // Wait to continue.
            Console.WriteLine("\nMain method complete. Press Enter");
            Console.ReadLine(); 

        }
    }
}
