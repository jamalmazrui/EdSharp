using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Security.Principal;
using System.Diagnostics;


namespace Recipe14_15
{
    class Program
    {
        static void Main(string[] args)
        {
            // check to see if the first argument is "elevated"
            if (args.Length > 0 && args[0] == "elevated")
            {
                Console.WriteLine("Started with command line argument");
                performElevatedTasks();
            }
            else
            {
                Console.WriteLine("Started without command line argument");
                performNormalTasks();
            }
        }

        static void performNormalTasks()
        {

            Console.WriteLine("Press return to perform elevated tasks");
            Console.ReadLine();
            // check to see if we have been started with elevated privileges
            if (checkElevatedPrivilege())
            {
                // we already have privileges - perform the tasks
                performElevatedTasks();
            }
            else
            {
                // we need to start an elevated instance
                startElevatedInstance();
            }
        }

        static void performElevatedTasks()
        {
            // check to see that we have elevated privileges
            if (checkElevatedPrivilege())
            {
                // perform the elevated task
                Console.WriteLine("Elevated tasks performed");
            }
            else
            {
                // we are not able to perform the elevated tasks
                Console.WriteLine("Cannot perform elevated tasks");
            }
            Console.WriteLine("Press return to exit");
            Console.ReadLine();
        }

        static bool checkElevatedPrivilege()
        {
            WindowsIdentity winIdentity = WindowsIdentity.GetCurrent();    
            WindowsPrincipal winPrincipal = new WindowsPrincipal(winIdentity);    
            return winPrincipal.IsInRole (WindowsBuiltInRole.Administrator);
        }

        static void startElevatedInstance()
        {
            ProcessStartInfo procstartinf = new ProcessStartInfo("Recipe14-15.exe");
            procstartinf.Arguments = "elevated";
            procstartinf.Verb = "runas";
            Process.Start(procstartinf).WaitForExit();
        }
    }
}
