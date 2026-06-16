using System;
using System.IO;
using System.Security.Principal;
using System.Security.Permissions;
using System.Runtime.InteropServices;

// Ensure the assembly has permission to execute unmanaged code
// and control the thread principal.
[assembly:SecurityPermission(SecurityAction.RequestMinimum,
    UnmanagedCode=true, ControlPrincipal=true)]

namespace Apress.VisualCSharpRecipes.Chapter11
{
    class Recipe11_12
    {
        // Define some constants for use with the LogonUser function.
        const int LOGON32_PROVIDER_DEFAULT = 0;
        const int LOGON32_LOGON_INTERACTIVE = 2;

        // Import the Win32 LogonUser function from advapi32.dll. Specify
        // "SetLastError = true" to correctly support access to Win32 error
        // codes.
        [DllImport("advapi32.dll", SetLastError=true, CharSet=CharSet.Unicode)]
        static extern bool LogonUser(string userName, string domain, 
            string password, int logonType, int logonProvider,
            ref IntPtr accessToken);

        public static void Main(string[] args) 
        {
            // Create a new IntPtr to hold the access token returned by the 
            // LogonUser function.
            IntPtr accessToken = IntPtr.Zero;
            
            // Call LogonUser to obtain an access token for the specified user.
            // The accessToken variable is passed to LogonUser by reference and 
            // will contain a reference to the Windows access token if 
            // LogonUser is successful.
            bool success = LogonUser(
                args[0],                    // username to log on.
                ".",                        // use the local account database.
                args[1],                    // user’s password.
                LOGON32_LOGON_INTERACTIVE,  // create an interactive login.
                LOGON32_PROVIDER_DEFAULT,   // use the default logon provider.
                ref accessToken             // receives access token handle.
            );
            
            // If the LogonUser return code is zero, an error has occurred. 
            // Display the error and exit.
            if (success)  
            {        
                Console.WriteLine("LogonUser returned error {0}", 
                    Marshal.GetLastWin32Error());
            } 
            else 
            {
                // Create a new WindowsIdentity from the Windows access token.
                WindowsIdentity identity = new WindowsIdentity(accessToken);
        
                // Display the active identity.
                Console.WriteLine("Identity before impersonation = {0}",
                    WindowsIdentity.GetCurrent().Name);

                // Impersonate the specified user, saving a reference to the 
                // returned WindowsImpersonationContext, which contains the 
                // information necessary to revert to the original user 
                // context.
                WindowsImpersonationContext impContext = 
                    identity.Impersonate();

                // Display the active identity.
                Console.WriteLine("Identity during impersonation = {0}",
                    WindowsIdentity.GetCurrent().Name);

                // *****************************************        
                // Perform actions as the impersonated user.
                // *****************************************      

                // Revert to the original Windows user using the 
                // WindowsImpersonationContext object.
                impContext.Undo();

                // Display the active identity.
                Console.WriteLine("Identity after impersonation  = {0}",
                    WindowsIdentity.GetCurrent().Name);

                // Wait to continue.
                Console.WriteLine("\nMain method complete. Press Enter.");
                Console.ReadLine();
            }
        }
    }
}