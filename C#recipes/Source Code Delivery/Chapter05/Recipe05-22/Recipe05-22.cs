using System;
using System.IO;
using System.Security.AccessControl;

namespace Apress.VisualCSharpRecipes.Chapter05
{
    static class Recipe05_22
    {
        static void Main(string[] args)
        {
            FileStream stream;
            string fileName;

            // Create a new file and assign full control to 'Everyone'.
            Console.WriteLine("Press any key to write a new file...");
            Console.ReadKey(true);

            fileName = Path.GetRandomFileName();
            using (stream = new FileStream(fileName, FileMode.Create))
            {
                // Do something.
            }
            Console.WriteLine("Created a new file " + fileName + ".");
            Console.WriteLine();


            // Deny 'Everyone' access to the file
            Console.WriteLine("Press any key to deny 'Everyone' " + 
                "access to the file...");
            Console.ReadKey(true);
            SetRule(fileName, "Everyone", 
                FileSystemRights.Read, AccessControlType.Deny);
            Console.WriteLine("Removed access rights of 'Everyone'.");
            Console.WriteLine();


            // Attempt to access file.
            Console.WriteLine("Press any key to attempt " + 
                "access to the file...");
            Console.ReadKey(true);

            try
            {
                stream = new FileStream(fileName, FileMode.Create);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Exception thrown: ");
                Console.WriteLine(ex.ToString());
            }
            finally
            {
                stream.Close();
                stream.Dispose();
            }

            // Wait to continue.
            Console.WriteLine(Environment.NewLine);
            Console.WriteLine("Main method complete. Press Enter.");
            Console.ReadLine();
        }

        static void AddRule(string filePath, string account, 
            FileSystemRights rights, AccessControlType controlType)
        {
            // Get a FileSecurity object that represents the 
            // current security settings.
            FileSecurity fSecurity = File.GetAccessControl(filePath);

            // Add the FileSystemAccessRule to the security settings. 
            fSecurity.AddAccessRule(new FileSystemAccessRule(account,
                rights, controlType));

            // Set the new access settings.
            File.SetAccessControl(filePath, fSecurity);
        }

        static void SetRule(string filePath, string account, 
            FileSystemRights rights, AccessControlType controlType)
        {
            // Get a FileSecurity object that represents the 
            // current security settings.
            FileSecurity fSecurity = File.GetAccessControl(filePath);

            // Add the FileSystemAccessRule to the security settings. 
            fSecurity.ResetAccessRule(new FileSystemAccessRule(account,
                rights, controlType));

            // Set the new access settings.
            File.SetAccessControl(filePath, fSecurity);
        }

    }
}
