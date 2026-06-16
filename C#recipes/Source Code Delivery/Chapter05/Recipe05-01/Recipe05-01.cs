using System;
using System.IO;

namespace Apress.VisualCSharpRecipes.Chapter05
{
    static class Recipe05_01
    {
        static void Main(string[] args)
        {
            if (args.Length == 0)
            {
                Console.WriteLine("Please supply a filename.");
                return;
            }

            // Display file information.
            FileInfo file = new FileInfo(args[0]);

            Console.WriteLine("Checking file: " + file.Name);
            Console.WriteLine("File exists: " + file.Exists.ToString());

            if (file.Exists)
            {
                Console.Write("File created: ");
                Console.WriteLine(file.CreationTime.ToString());
                Console.Write("File last updated: ");
                Console.WriteLine(file.LastWriteTime.ToString());
                Console.Write("File last accessed: ");
                Console.WriteLine(file.LastAccessTime.ToString());
                Console.Write("File size (bytes): ");
                Console.WriteLine(file.Length.ToString());
                Console.Write("File attribute list: ");
                Console.WriteLine(file.Attributes.ToString());
            }
            Console.WriteLine();

            // Display directory information.
            DirectoryInfo dir = file.Directory;

            Console.WriteLine("Checking directory: " + dir.Name);
            Console.WriteLine("In directory: " + dir.Parent.Name);
            Console.Write("Directory exists: ");
            Console.WriteLine(dir.Exists.ToString());

            if (dir.Exists)
            {
                Console.Write("Directory created: ");
                Console.WriteLine(dir.CreationTime.ToString());
                Console.Write("Directory last updated: ");
                Console.WriteLine(dir.LastWriteTime.ToString());
                Console.Write("Directory last accessed: ");
                Console.WriteLine(dir.LastAccessTime.ToString());
                Console.Write("Directory attribute list: ");
                Console.WriteLine(dir.Attributes.ToString());
                Console.WriteLine("Directory contains: " +
                  dir.GetFiles().Length.ToString() + " files");
            }
            Console.WriteLine();

            // Display drive information.
            DriveInfo drv = new DriveInfo(file.FullName);

            Console.Write("Drive: ");
            Console.WriteLine(drv.Name);
            
            if (drv.IsReady)
            {
                Console.Write("Drive type: ");
                Console.WriteLine(drv.DriveType.ToString());
                Console.Write("Drive format: ");
                Console.WriteLine(drv.DriveFormat.ToString());
                Console.Write("Drive free space: ");
                Console.WriteLine(drv.AvailableFreeSpace.ToString());
            }

            // Wait to continue.
            Console.WriteLine(Environment.NewLine);
            Console.WriteLine("Main method complete. Press Enter.");
            Console.ReadLine();
        }
    }
}
