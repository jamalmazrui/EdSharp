using System;
using System.Security.Principal;

namespace Apress.VisualCSharpRecipes.Chapter11
{
    class Recipe11_10
    {
        public static void Main (string[] args) 
        {
            if (args.Length == 0)
            {
                Console.WriteLine("Please provide groups to check as command line arguments");
            }

            // Obtain a WindowsIdentity object representing the currently
            // logged on Windows user.
            WindowsIdentity identity = WindowsIdentity.GetCurrent();

            // Create a WindowsPrincipal object that represents the security
            // capabilities of the specified WindowsIdentity; in this case
            // the Windows groups to which the current user belongs.
            WindowsPrincipal principal = new WindowsPrincipal(identity);

            // Iterate through the group names specified as command-line 
            // arguments and test to see if the current user is a member of 
            // each one.
            foreach (string role in args) 
            {
                Console.WriteLine("Is {0} a member of {1}? = {2}", 
                    identity.Name, role, principal.IsInRole(role));
            }

            // Wait to continue.
            Console.WriteLine("\nMain method complete. Press Enter.");
            Console.ReadLine();
        }
    }
}