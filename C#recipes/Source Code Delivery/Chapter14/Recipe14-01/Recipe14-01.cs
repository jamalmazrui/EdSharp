using System;

namespace Apress.VisualCSharpRecipes.Chapter14
{
    class Recipe14_01
    {
        public static void Main() 
        {
            // Command line.
            Console.WriteLine("Command line : " + Environment.CommandLine);

            // OS and CLR version information.
            Console.WriteLine(Environment.NewLine);
            Console.WriteLine("OS PlatformID : " + 
                Environment.OSVersion.Platform);
            Console.WriteLine("OS Major Version : " + 
                Environment.OSVersion.Version.Major);
            Console.WriteLine("OS Minor Version : " + 
                Environment.OSVersion.Version.Minor);
            Console.WriteLine("CLR Version : " + Environment.Version);

            // User, machine, and domain name information.
            Console.WriteLine(Environment.NewLine);
            Console.WriteLine("User Name : " + Environment.UserName);
            Console.WriteLine("Domain Name : " + Environment.UserDomainName);
            Console.WriteLine("Machine name : " + Environment.MachineName);

            // Other environment information.
            Console.WriteLine(Environment.NewLine);
            Console.WriteLine("Is interactive? : " 
                + Environment.UserInteractive);
            Console.WriteLine("Shutting down? : " 
                + Environment.HasShutdownStarted);
            Console.WriteLine("Ticks since startup : " 
                + Environment.TickCount);

            // Display the names of all logical drives.
            Console.WriteLine(Environment.NewLine);
            foreach (string s in Environment.GetLogicalDrives()) 
            {
                Console.WriteLine("Logical drive : " + s);
            }

            // Standard folder information.
            Console.WriteLine(Environment.NewLine);
            Console.WriteLine("Current folder : " 
                + Environment.CurrentDirectory);
            Console.WriteLine("System folder : " 
                + Environment.SystemDirectory);

            // Enumerate all special folders and display them.
            Console.WriteLine(Environment.NewLine);
            foreach (Environment.SpecialFolder s in 
                Enum.GetValues(typeof(Environment.SpecialFolder))) 
            {
                Console.WriteLine("{0} folder : {1}", 
                    s, Environment.GetFolderPath(s));
            }

            // Wait to continue.
            Console.WriteLine(Environment.NewLine);
            Console.WriteLine("Main method complete. Press Enter.");
            Console.ReadLine();
        }
    }
}
