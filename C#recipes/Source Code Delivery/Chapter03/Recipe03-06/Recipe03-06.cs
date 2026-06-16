using System;

namespace Apress.VisualCSharpRecipes.Chapter03
{
    class Recipe03_06
    {
        public static void Main(string[] args)
        {
            // For the purpose of this example, if this assembly is executing
            // in an AppDomain with the friendly name "NewAppDomain", do not 
            // create a new AppDomain. This avoids an infinite loop of 
            // AppDomain creation.

            if (AppDomain.CurrentDomain.FriendlyName != "NewAppDomain")
            {
                // Create a new application domain.
                AppDomain domain = AppDomain.CreateDomain("NewAppDomain");

                // Execute this assembly in the new application domain and
                // pass the array of command-line arguments.
                domain.ExecuteAssembly("Recipe03-06.exe", args);
            }

            // Display the command-line arguments to the screen prefixed with
            // the friendly name of the AppDomain.
            foreach (string s in args)
            {
                Console.WriteLine(AppDomain.CurrentDomain.FriendlyName + " : " + s);
            }

            // Wait to continue.
            if (AppDomain.CurrentDomain.FriendlyName != "NewAppDomain")
            {
                Console.WriteLine("\nMain method complete. Press Enter.");
                Console.ReadLine();
            }
        }
    }
}
