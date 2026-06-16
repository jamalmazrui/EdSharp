using System;
using System.Threading;

namespace Apress.VisualCSharpRecipes.Chapter04
{
    class Recipe04_09
    {
        // Boolean to signal that the second thread should terminate.
        static bool terminate = false;

        // A utility method for displaying useful trace information to the 
        // console along with details of the current thread.
        private static void TraceMsg(string msg)
        {
            Console.WriteLine("[{0,3}] - {1} : {2}",
                Thread.CurrentThread.ManagedThreadId,
                DateTime.Now.ToString("HH:mm:ss.ffff"), msg);
        }

        // Declare the method that will be executed on the separate thread.
        // In a loop the method waits to obtain a Mutex before displaying a 
        // message to the console, then waits 1 second before releasing the 
        // Mutex.
        private static void DisplayMessage()
        {
            // Obtain a handle to the Mutex with the name "MutexExample".
            // Don't attempt to take ownership immediately.
            using (Mutex mutex = new Mutex(false, "MutexExample"))
            {
                TraceMsg("Thread started.");

                while (!terminate)
                {
                    // Wait on the Mutex.
                    mutex.WaitOne();

                    TraceMsg("Thread owns the Mutex.");

                    Thread.Sleep(1000);

                    TraceMsg("Thread releasing the Mutex.");

                    // Release the Mutex.
                    mutex.ReleaseMutex();

                    // Sleep a little to give another thread a good chance of
                    // acquiring the Mutex.
                    Thread.Sleep(100);
                }

                TraceMsg("Thread terminating.");
            }
        }

        public static void Main()
        {
            // Create a new Mutex with the name "MutexExample".
            using (Mutex mutex = new Mutex(false, "MutexExample"))
            {
                TraceMsg("Starting threads -- press Enter to terminate.");

                // Create and start 3 new threads running the 
                // DisplayMesssage method.
                Thread trd1 = new Thread(DisplayMessage);
                Thread trd2 = new Thread(DisplayMessage);
                Thread trd3 = new Thread(DisplayMessage);
                trd1.Start();
                trd2.Start();
                trd3.Start();

                // Wait for Enter to be pressed.
                Console.ReadLine();

                // Terminate the DisplayMessage threads and wait for them to 
                // complete before disposing of the Mutex.
                terminate = true;
                trd1.Join(5000);
                trd2.Join(5000);
                trd3.Join(5000);
            }

            // Wait to continue.
            Console.WriteLine(Environment.NewLine);
            Console.WriteLine("Main method complete. Press Enter.");
            Console.ReadLine();
        }
    }
}
