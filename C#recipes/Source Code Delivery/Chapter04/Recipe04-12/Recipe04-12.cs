using System;
using System.Threading;

namespace Apress.VisualCSharpRecipes.Chapter04
{
    class Recipe04_12
    {
        private static void DisplayMessage()
        {
            // Display a message to the console 5 times.
            for (int count = 0; count < 5; count++)
            {
                Console.WriteLine("{0} : DisplayMessage thread",
                    DateTime.Now.ToString("HH:mm:ss.ffff"));

                // Sleep for 1 second.
                Thread.Sleep(1000);
            }
        }

        public static void Main()
        {
            // Create a new Thread to run the DisplayMessage method.
            Thread thread = new Thread(DisplayMessage);

            Console.WriteLine("{0} : Starting DisplayMessage thread.",
                DateTime.Now.ToString("HH:mm:ss.ffff"));

            // Start the DisplayMessage thread.
            thread.Start();

            // Block until the DisplayMessage thread finishes, or timeout after
            // 2 seconds.
            if (!thread.Join(2000))
            {
                Console.WriteLine("{0} : Join timed out !!",
                    DateTime.Now.ToString("HH:mm:ss.ffff"));
            }

            // print out the thread status
            Console.WriteLine("Thread alive: {0}", thread.IsAlive);

            // Block again until the DisplayMessage thread finishes with no timeout.
            thread.Join();

            // print out the thread status
            Console.WriteLine("Thread alive: {0}", thread.IsAlive);

            // Wait to continue.
            Console.WriteLine("Main method complete. Press Enter.");
            Console.ReadLine();
        }
    }
}
