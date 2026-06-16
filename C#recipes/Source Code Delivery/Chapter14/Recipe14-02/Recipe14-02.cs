using System;
using System.Collections;

namespace Apress.VisualCSharpRecipes.Chapter14
{
    class Recipe14_02
    {
        public static void Main()
        {
            // Retrieve a named environment variable.
            Console.WriteLine("Path = " +
                Environment.GetEnvironmentVariable("Path"));
            Console.WriteLine(Environment.NewLine);

            // Substitute the value of named environment variables.
            Console.WriteLine(Environment.ExpandEnvironmentVariables(
                    "The Path on %computername% is %Path%"));
            Console.WriteLine(Environment.NewLine);

            // Retrieve all environment variables targeted at the process and  
            // display the values of all that begin with the letter ’U’.
            IDictionary vars = Environment.GetEnvironmentVariables(
                EnvironmentVariableTarget.Process);

            foreach (string s in vars.Keys)
            {
                if (s.ToUpper().StartsWith("U"))
                {
                    Console.WriteLine(s + " = " + vars[s]);
                }
            }

            // Wait to continue.
            Console.WriteLine(Environment.NewLine);
            Console.WriteLine("Main method complete. Press Enter.");
            Console.ReadLine();
        }
    }
}
