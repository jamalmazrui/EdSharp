using System;
using System.Threading;

namespace Apress.VisualCSharpRecipes.Chapter04
{
    class Recipe04_13
    {
        private static void DisplayMessage()
        {
            try
            {
                while (true)
                {
                    // Display a message to the console.
                    Console.WriteLine("{0} : DisplayMessage thread active",
                        DateTime.Now.ToString("HH:mm:ss.ffff"));

                    // Sleep for 1 second.
                    Thread.Sleep(1000);
                }
            }
            catch (ThreadAbortException ex)
            {
                // Display a message to the console.
                Console.WriteLine("{0} : DisplayMessage thread terminating - {1}",
                    DateTime.Now.ToString("HH:mm:ss.ffff"),
                    (string)ex.ExceptionState);

                // Call Thread.ResetAbort here to cancel the abort request.
            }

            // This code is never executed unless Thread.ResetAbort
            // is called in the previous catch block.
            Console.WriteLine("{0} : Nothing is called after the catch block",
                DateTime.Now.ToString("HH:mm:ss.ffff"));
        }

        public static void Main()
        {
            // Create a new Thread to run the DisplayMessage method.
            Thread thread = new Thread(DisplayMessage);

            Console.WriteLine("{0} : Starting DisplayMessage thread" +
                " - press Enter to terminate.",
                DateTime.Now.ToString("HH:mm:ss.ffff"));

            // Start the DisplayMessage thread.
            thread.Start();

            // Wait until Enter is pressed and terminate the thread.
            Console.ReadLine();

            thread.Abort("User pressed Enter");

            // Block again until the DisplayMessage thread finishes.
            thread.Join();

            // Wait to continue.
            Console.WriteLine("Main method complete. Press Enter.");
            Console.ReadLine();
        }
    }
}

