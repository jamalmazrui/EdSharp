using System;
using System.IO;
using System.IO.IsolatedStorage;

namespace Apress.VisualCSharpRecipes.Chapter05
{
    static class Recipe05_18
    {
        static void Main()
        {
            // Create the store for the current user.
            using (IsolatedStorageFile store = 
                IsolatedStorageFile.GetUserStoreForAssembly())
            {
                // Create a folder in the root of the isolated store.
                store.CreateDirectory("MyFolder");

                // Create a file in the isolated store.
                using (Stream fs = new IsolatedStorageFileStream(
                  "MyFile.txt", FileMode.Create, store))
                {
                    StreamWriter w = new StreamWriter(fs);

                    // You can now write to the file as normal.
                    w.WriteLine("Test");
                    w.Flush();
                }

                Console.WriteLine("Current size: " + 
                    store.UsedSize.ToString());
                Console.WriteLine("Scope: " + store.Scope.ToString());

                Console.WriteLine("Contained files include:");
                string[] files = store.GetFileNames("*.*");
                foreach (string file in files)
                {
                    Console.WriteLine(file);
                }
            }
            
            // Wait to continue.
            Console.WriteLine(Environment.NewLine);
            Console.WriteLine("Main method complete. Press Enter.");
            Console.ReadLine();
        }
    }
}
