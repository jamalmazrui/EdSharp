using System;
using System.Threading;

namespace Apress.VisualCSharpRecipes.Chapter04
{
    class Recipe04_10
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
        // In a loop the method waits to obtain a Semaphore before displaying a 
        // message to the console, then waits 1 second before releasing the 
        // Semaphore.
        private static void DisplayMessage()
        {
            // Obtain a handle to the Semaphore with the name "SemaphoreExample".
            using (Semaphore sem = Semaphore.OpenExisting("SemaphoreExample"))
            {
                TraceMsg("Thread started.");

                while (!terminate)
                {
                    // Wait on the Semaphore.
                    sem.WaitOne();

                    TraceMsg("Thread owns the Semaphore.");

                    Thread.Sleep(1000);

                    TraceMsg("Thread releasing the Semaphore.");

                    // Release the Semaphore.
                    sem.Release();

                    // Sleep a little to give another thread a good chance of
                    // acquiring the Semaphore.
                    Thread.Sleep(100);
                }

                TraceMsg("Thread terminating.");
            }
        }

        public static void Main()
        {
            // Create a new Semaphore with the name "SemaphoreExample". The
            // Semaphore can be owned by up to 2 threads at the same time.
            using (Semaphore sem = new Semaphore(2,2,"SemaphoreExample"))
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
                // complete before disposing of the Semaphore.
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
