using System;

namespace Apress.VisualCSharpRecipes.Chapter03
{
    class Recipe03_01
    {
        public static void Main()
        {
            // Instantiate an AppDomainSetup object.
            AppDomainSetup setupInfo = new AppDomainSetup();

            // Configure the application domain setup information.
            setupInfo.ApplicationBase = @"C:\MyRootDirectory";
            setupInfo.ConfigurationFile = "MyApp.config";
            setupInfo.PrivateBinPath = "bin;plugins;external";

            // Create a new application domain passing null as the evidence 
            // argument. Remember to save a reference to the new AppDomain as 
            // this cannot be retrieved any other way.
            AppDomain newDomain = 
                AppDomain.CreateDomain("My New AppDomain", null, setupInfo);

            // Wait to continue.
            Console.WriteLine("\nMain method complete. Press Enter.");
            Console.ReadLine();
        }
    }
}
