using System;
using System.IO;
using IWshRuntimeLibrary;

namespace Apress.VisualCSharpRecipes.Chapter14
{
    class Recipe14_08
    {
        public static void CreateShortcut(string destination)
        {
            // Create a WshShell instance through which to access the 
            // functionality of the Windows shell.
            WshShell wshShell = new WshShell();

            // Assemble a fully qualified name that places the Notepad.lnk
            // file in the specified destination folder. You could use the 
            // System.Environment.GetFolderPath method to obtain a path, but 
            // the WshShell.SpecialFolders method provides access to a wider 
            // range of folders. You need to create a temporary object reference 
            // to the destination string to satisfy the requirements of the 
            // Item method signature.
            object destFolder = (object)destination;
            string fileName = Path.Combine(
                (string)wshShell.SpecialFolders.Item(ref destFolder),
                "Notepad.lnk"
            );

            // Create the shortcut object. Nothing is created in the 
            // destination folder until the shortcut is saved.
            IWshShortcut shortcut =
                (IWshShortcut)wshShell.CreateShortcut(fileName);

            // Configure the fully qualified name to the executable.
            // Use the Environment class for simplicity.
            shortcut.TargetPath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.System),
                "notepad.exe"
            );

            // Set the working directory to the Personal (My Documents) folder.
            shortcut.WorkingDirectory =
                Environment.GetFolderPath(Environment.SpecialFolder.Personal);

            // Provide a description for the shortcut.
            shortcut.Description = "Notepad Text Editor";

            // Assign a hotkey to the shortcut.
            shortcut.Hotkey = "CTRL+ALT+N";

            // Configure Notepad to always start maximized.
            shortcut.WindowStyle = 3;

            // Configure the shortcut to display the first icon in Notepad.exe.
            shortcut.IconLocation = "notepad.exe, 0";

            // Save the configured shortcut file.
            shortcut.Save();
        }

        public static void Main()
        {
            // Create the Notepad shortcut on the desktop.
            CreateShortcut("Desktop");

            // Create the Notepad shortcut on the Windows Start menu of 
            // the current user.
            CreateShortcut("StartMenu");

            // Wait to continue.
            Console.WriteLine(Environment.NewLine);
            Console.WriteLine("Main method complete. Press Enter.");
            Console.ReadLine();
        }
    }
}
