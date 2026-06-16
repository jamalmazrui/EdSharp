using System;
using System.Security;
using System.Diagnostics;

namespace Apress.VisualCSharpRecipes.Chapter11
{
    class Recipe11_18
    {
        public static SecureString ReadString()
        {
            // Create a new emtpty SecureString.
            SecureString str = new SecureString();

            // Read the string from the console one
            // character at a time without displaying it. 
            ConsoleKeyInfo nextChar = Console.ReadKey(true);

            // Read characters until Enter is pressed.
            while (nextChar.Key != ConsoleKey.Enter)
            {
                if (nextChar.Key == ConsoleKey.Backspace)
                {
                    if (str.Length > 0)
                    {
                        // Backspace pressed, remove the last character.
                        str.RemoveAt(str.Length - 1);

                        Console.Write(nextChar.KeyChar);
                        Console.Write(" ");
                        Console.Write(nextChar.KeyChar);
                    }
                    else
                    {
                        Console.Beep();
                    }
                }
                else
                {
                    // Append the character to the SecureString and
                    // display a masked character.
                    str.AppendChar(nextChar.KeyChar);
                    Console.Write("*");
                }
                    
                // Read the next character.
                nextChar = Console.ReadKey(true);
            }

            // String entry finished. Make it read-only.
            str.MakeReadOnly();
            return str;
        }

        public static void Main()
        {
            string user = "";

            // Get the username under which Notepad.exe will be run.
            Console.Write("Enter the user name: ");
            user = Console.ReadLine();

            // Get the user's password as a SecureString.
            Console.Write("Enter the user's password: ");
            using (SecureString pword = ReadString())
            {
                // Start Notepad as the specified user.
                ProcessStartInfo startInfo = new ProcessStartInfo();

                startInfo.FileName = "notepad.exe";
                startInfo.UserName = user;
                startInfo.Password = pword;
                startInfo.UseShellExecute = false;

                // Create a new Process object.
                using (Process process = new Process())
                {
                    // Assign the ProcessStartInfo to the Process object.
                    process.StartInfo = startInfo;

                    try
                    {
                        // Start the new process.
                        process.Start();
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("\n\nCould not start Notepad process.");
                        Console.WriteLine(ex);
                    }
                }
            }

            // Wait to continue.
            Console.WriteLine("\n\nMain method complete. Press Enter.");
            Console.ReadLine();
        }
    }
}
