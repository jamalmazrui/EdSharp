using System;
using System.Reflection;
using System.Collections;
using System.Security.Policy;

namespace Apress.VisualCSharpRecipes.Chapter11
{
    public class Recipe11_09 
    {
        public static void Main(string[] args) 
        {
            // Load the specified assembly.
            Assembly a = Assembly.LoadFrom(args[0]);

            // Get the Evidence collection from the 
            // loaded assembly.
            Evidence e = a.Evidence;

            // Display the Host Evidence.
            IEnumerator x = e.GetHostEnumerator();
            Console.WriteLine("HOST EVIDENCE COLLECTION:");
            while(x.MoveNext()) 
            {
                Console.WriteLine(x.Current.ToString());
                Console.WriteLine("Press Enter to see next evidence.");
                Console.ReadLine();
            }

            // Display the Assembly Evidence.
            x = e.GetAssemblyEnumerator();
            Console.WriteLine("ASSEMBLY EVIDENCE COLLECTION:");
            while(x.MoveNext()) 
            {
                Console.WriteLine(x.Current.ToString());
                Console.WriteLine("Press Enter to see next evidence.");
                Console.ReadLine();
            }

            // Wait to continue.
            Console.WriteLine("Main method complete. Press Enter.");
            Console.ReadLine();
        }
    }
}