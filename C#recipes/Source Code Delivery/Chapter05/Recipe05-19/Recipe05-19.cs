using System;
using System.IO;
using System.Windows.Forms;

namespace Apress.VisualCSharpRecipes.Chapter05
{
    static class Recipe05_19
    {
        static void Main()
        {
            // Configure the FileSystemWatcher.
            using (FileSystemWatcher watch = new FileSystemWatcher())
            {
                watch.Path = Application.StartupPath;
                watch.Filter = "*.*";
                watch.IncludeSubdirectories = true;

                // Attach the event handler.
                watch.Created += new FileSystemEventHandler(OnCreatedOrDeleted);
                watch.Deleted += new FileSystemEventHandler(OnCreatedOrDeleted);
                watch.EnableRaisingEvents = true;

                Console.WriteLine("Press Enter to create a file.");
                Console.ReadLine();

                if (File.Exists("test.bin"))
                {
                    File.Delete("test.bin");
                }

                // Create test.bin.
                using (FileStream fs = new FileStream("test.bin", FileMode.Create))
                {
                    // Do something.
                }

                Console.WriteLine("Press Enter to terminate the application.");
                Console.ReadLine();
            }
            
            // Wait to continue.
            Console.WriteLine(Environment.NewLine);
            Console.WriteLine("Main method complete. Press Enter.");
            Console.ReadLine();
        }

        // Fires when a new file is created in the directory being monitored.
        private static void OnCreatedOrDeleted(object sender,
          FileSystemEventArgs e)
        {
            // Display the notification information.
            Console.WriteLine("\tNOTIFICATION: " + e.FullPath +
              "' was " + e.ChangeType.ToString());
            Console.WriteLine();
        }
    }
}
