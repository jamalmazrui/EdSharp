using System;
using System.Threading;

namespace Apress.VisualCSharpRecipes.Chapter04
{
    class Recipe04_17
    {
        public static void Main()
        {
            // A boolean that indicates whether this application has
            // initial ownership of the Mutex.
            bool ownsMutex;

            // Attempt to create and take ownership of a Mutex named
            // MutexExample.
            using (Mutex mutex =
                       new Mutex(true, "MutexExample", out ownsMutex))
            {
                // If the application owns the Mutex it can continue to execute;
                // otherwise, the application should exit. 
                if (ownsMutex)
                {
                    Console.WriteLine("This application currently owns the" +
                        " mutex named MutexExample. Additional instances of" +
                        " this application will not run until you release" +
                        " the mutex by pressing Enter.");

                    Console.ReadLine();

                    // Release the mutex
                    mutex.ReleaseMutex();
                }
                else
                {
                    Console.WriteLine("Another instance of this application " +
                        " already owns the mutex named MutexExample. This" +
                        " instance of the application will terminate.");
                }
            }

            // Wait to continue.
            Console.WriteLine("Main method complete. Press Enter.");
            Console.ReadLine();
        }
    }
}
