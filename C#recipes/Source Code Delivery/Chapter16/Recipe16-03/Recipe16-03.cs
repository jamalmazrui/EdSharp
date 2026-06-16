using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Apress.VisualCSharpRecipes.Chapter16
{
    class Recipe16_03
    {
        static void Main(string[] args)
        {
            IList<object> mixedData = createData();
            IEnumerable <string> stringData = mixedData.OfType<string>();
            foreach (string str in stringData)
            {
                Console.WriteLine(str);
            }

            // Wait to continue.
            Console.WriteLine("\nMain method complete. Press Enter");
            Console.ReadLine(); 
        }

        static IList<object> createData()
        {
            return new List<object>()
            {
                "this is a string",
                23,
                9.2,
                "this is another string"
            };
        }
    }
}
