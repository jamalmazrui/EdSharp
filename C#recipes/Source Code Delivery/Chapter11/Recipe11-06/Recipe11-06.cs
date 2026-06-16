using System;
using System.Net;
using System.Security.Permissions;

// Minimum permission request for SocketPermission.
[assembly: SocketPermission(SecurityAction.RequestMinimum,
    Unrestricted = true)]

// Optional permission request for IsolatedStorageFilePermission.
[assembly: IsolatedStorageFilePermission(SecurityAction.RequestOptional,
    Unrestricted = true)]

// Refuse request for ReflectionPermission.
[assembly: ReflectionPermission(SecurityAction.RequestRefuse,
    Unrestricted = true)]

namespace Apress.VisualCSharpRecipes.Chapter11
{
    class Recipe11_06
    {
        public static void Main()
        {
            // Create and configure a FileIOPermission object that represents 
            // write access to the C:\Data folder.
            FileIOPermission fileIOPerm =
                new FileIOPermission(FileIOPermissionAccess.Write, @"C:\Data");

            // Make the demand.
            fileIOPerm.Demand();

            // Do something...

            // Wait to continue.
            Console.WriteLine("Main method complete. Press Enter.");
            Console.ReadLine();
        }
    }
}