using System;
using Microsoft.Win32;

namespace Apress.VisualCSharpRecipes.Chapter14
{
    class Recipe14_04
    {
        public static void Main(String[] args)
        {
            // Variables to hold usage information read from registry.
            string lastUser;
            string lastRun;
            int runCount;

            // Read the name of the last user to run the application from the 
            // registry. This is stored as the default value of the key and is
            // accessed by not specifying a value name. Cast the returned Object 
            // to a string.
            lastUser = (string)Registry.GetValue(
                @"HKEY_CURRENT_USER\Software\Apress\Visual C# Recipes",
                "", "Nobody");
 
            // If lastUser is null, it means that the specified registry key
            // does not exist.
            if (lastUser == null)
            {
                // Set initial values for the usage information.
                lastUser = "Nobody";
                lastRun = "Never";
                runCount = 0;
            }
            else
            {
                // Read the last run date and specify a default value of
                // "Never". Cast the returned Object to string.
                lastRun = (string)Registry.GetValue(
                    @"HKEY_CURRENT_USER\Software\Apress\Visual C# Recipes",
                    "LastRun", "Never");

                // Read the run count value and specify a default value of 
                // 0 (zero). Cast the Object to Int32, and assign to an int.
                runCount = (Int32)Registry.GetValue(
                    @"HKEY_CURRENT_USER\Software\Apress\Visual C# Recipes",
                    "RunCount", 0);
            }

            // Display the usage information.
            Console.WriteLine("Last user name: " + lastUser);
            Console.WriteLine("Last run date/time: " + lastRun);
            Console.WriteLine("Previous executions: " + runCount);

            // Update the usage information. It doesn't matter if the registry 
            // key exists or not, SetValue will automatically create it.

            // Update the "last user" information with the current username.
            // Specify that this should be stored as the default value
            // for the key by using an empty string as the value name.
            Registry.SetValue(
                @"HKEY_CURRENT_USER\Software\Apress\Visual C# Recipes",
                "", Environment.UserName, RegistryValueKind.String);

            // Update the "last run" information with the current date and time.
            // Specify that this should be stored as a string value in the 
            // registry.
            Registry.SetValue(
                @"HKEY_CURRENT_USER\Software\Apress\Visual C# Recipes",
                "LastRun", DateTime.Now.ToString(), RegistryValueKind.String);

            // Update the usage count information. Specify that this should 
            // be stored as an integer value in the registry.
            Registry.SetValue(
                @"HKEY_CURRENT_USER\Software\Apress\Visual C# Recipes",
                "RunCount", ++runCount, RegistryValueKind.DWord);

            // Wait to continue.
            Console.WriteLine(Environment.NewLine);
            Console.WriteLine("Main method complete. Press Enter.");
            Console.ReadLine();
        }
    }
}
