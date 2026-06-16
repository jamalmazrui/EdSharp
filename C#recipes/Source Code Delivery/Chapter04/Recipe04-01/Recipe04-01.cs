using System;
using System.Threading;

namespace Apress.VisualCSharpRecipes.Chapter04
{
    class Recipe04_01
    {

        // A private class used to pass data to the DisplayMessage method when it is
        // executed using the thread pool.
        private class MessageInfo
        {
            private int iterations;
            private string message;

            // A constructor that takes configuration settings for the thread.
            public MessageInfo(int iterations, string message)
            {
                this.iterations = iterations;
                this.message = message;
            }

            // Properties to retrieve configuration settings.
            public int Iterations { get { return iterations; } }
            public string Message { get { return message; } }
        }

        // A method that conforms to the System.Threading.WaitCallback delegate
        // signature. Displays a message to the console.
        public static void DisplayMessage(object state)
        {
            // Safely cast the state argument to a MessageInfo object.
            MessageInfo config = state as MessageInfo;

            // If the config argument is null, no arguments were passed to
            // the ThreadPool.QueueUserWorkItem method, use default values.
            if (config == null)
            {
                // Display a fixed message to the console 3 times.
                for (int count = 0; count < 3; count++)
                {
                    Console.WriteLine("A thread pool example.");

                    // Sleep for the purpose of demonstration. Avoid sleeping
                    // on thread-pool threads in real applications.
                    Thread.Sleep(1000);
                }

            }
            else
            {
                // Display the specified message the specified number of times.
                for (int count = 0; count < config.Iterations; count++)
                {
                    Console.WriteLine(config.Message);

                    // Sleep for the purpose of demonstration. Avoid sleeping
                    // on thread-pool threads in real applications.
                    Thread.Sleep(1000);
                }
            }
        }

        public static void Main()
        {
            // Execute DisplayMessage using the thread pool and no arguments.
            ThreadPool.QueueUserWorkItem(DisplayMessage);

            // Create a MessageInfo object to pass to the DisplayMessage method.
            MessageInfo info = new MessageInfo(5, 
                "A thread pool example with arguments.");

            // set the max number of threads
            ThreadPool.SetMaxThreads(2, 2);

            // Execute DisplayMessage using the thread pool and providing an 
            // argument. 
            ThreadPool.QueueUserWorkItem(DisplayMessage, info);

            // Wait to continue.
            Console.WriteLine("Main method complete. Press Enter.");
            Console.ReadLine();
        }
    }
}