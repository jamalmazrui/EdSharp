using System;
using System.Net;
using System.Security.Permissions;

// Permission request for a SocketPermission that allows the code to open 
// a TCP connection to the specified host and port.
[assembly:SocketPermission(SecurityAction.RequestMinimum, 
    Access = "Connect", Host = "www.fabrikam.com",
    Port = "3538", Transport = "Tcp")]

// Permission request for the UnmanagedCode element of SecurityPermission, 
// which controls the code’s ability to execute unmanaged code.
[assembly:SecurityPermission(SecurityAction.RequestMinimum, 
    UnmanagedCode = true)]

namespace Apress.VisualCSharpRecipes.Chapter11
{
    class Recipe11_04
    {
        public static void Main()
        {        
            // Do something...

            // Wait to continue.
            Console.WriteLine("Main method complete. Press Enter.");
            Console.ReadLine();
        }
    }
}