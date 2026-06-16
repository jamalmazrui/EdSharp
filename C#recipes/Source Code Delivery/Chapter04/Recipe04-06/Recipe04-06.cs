using System;
using System.Threading;

namespace Apress.VisualCSharpRecipes.Chapter04
{
    class Recipe04_06
    {

        // A utility method for displaying useful trace information to the 
        // console along with details of the current thread.
        private static void TraceMsg(string msg)
        {
            Console.WriteLine("[{0,3}] - {1} : {2}",
                Thread.CurrentThread.ManagedThreadId,
                DateTime.Now.ToString("HH:mm:ss.ffff"), msg);
        }

        // A private class used to pass initialization data to a new thread.
        private class ThreadStartData
        {
            public ThreadStartData(int iterations, string message, int delay)
            {
                this.iterations = iterations;
                this.message = message;
                this.delay = delay;
            }
            
            // Member variables hold initialization data for a new thread.
            private readonly int iterations;
            private readonly string message;
            private readonly int delay;

            // Properties provide read-only access to initialization data.
            public int Iterations { get { return iterations; } }
            public string Message { get { return message; } }
            public int Delay { get { return delay; } }
        }

        // Declare the method that will be executed in its own thread. The 
        // method displays a message to the console a specified number of 
        // times, sleeping between each message for a specified duration.
        private static void DisplayMessage(object config)
        {
            ThreadStartData data = config as ThreadStartData;

            if (data != null)
            {
                for (int count = 0; count < data.Iterations; count++)
                {
                    TraceMsg(data.Message);

                    // Sleep for the specified period.
                    Thread.Sleep(data.Delay);
                }
            }
            else
            {
                TraceMsg("Invalid thread configuration.");
            }
        }

        public static void Main() 
        {
            // Create a new Thread object specifying DisplayMessage
            // as the method it will execute.
            Thread thread = new Thread(DisplayMessage);

            // make this a foreground thread - this is the 
            // default - call used for example purposes
            thread.IsBackground = false;

            // Create a new ThreadStartData object to configure the thread.
            ThreadStartData config =
                new ThreadStartData(5, "A thread example.", 500);

            TraceMsg("Starting new thread.");

            // Start the new thread and pass the ThreadStartData object
            // containing the initialization data.
            thread.Start(config);

            // Continue with other processing.
            for (int count = 0; count < 13; count++) 
            {
                TraceMsg("Main thread continuing processing...");
                Thread.Sleep(200);
            }

            // Wait to continue.
            Console.WriteLine(Environment.NewLine);
            Console.WriteLine("Main method complete. Press Enter.");
            Console.ReadLine();
        }
    }
}