using System;
using System.Threading;
using System.Collections.Generic;

namespace Apress.VisualCSharpRecipes.Chapter04
{
    class Recipe04_07
    {
        // Declare an object for synchronization of access to the Console. 
        // A static object is used because we are using it in static methods.
        private static object consoleGate = new Object();

        // Declare a Queue to represent the work queue.
        private static Queue<string> workQueue = new Queue<string>();

        // Declare a flag to indicate to activated threads that they should
        // terminate and not process more work items.
        private static bool processWorkItems = true;

        // A utility method for displaying useful trace information to the 
        // console along with details of the current thread.
        private static void TraceMsg(string msg)
        {
            lock (consoleGate)
            {
                Console.WriteLine("[{0,3}/{1}] - {2} : {3}",
                    Thread.CurrentThread.ManagedThreadId,
                    Thread.CurrentThread.IsThreadPoolThread ? "pool" : "fore",
                    DateTime.Now.ToString("HH:mm:ss.ffff"), msg);
            }
        }

        // Declare the method that will be executed by each thread to process
        // items from the work queue.
        private static void ProcessWorkItems() 
        {
            // A local variable to hold the work item taken from the work queue.
            string workItem = null;

            TraceMsg("Thread started, processing items from queue...");

            // Process items from the work queue until termination is signaled.
            while (processWorkItems)
            {
                // Obtain the lock on the work queue;
                Monitor.Enter(workQueue);

                try
                {
                    // Pop the next work item and process it, or wait if none
                    // are available.
                    if (workQueue.Count == 0)
                    {
                        TraceMsg("No work items, waiting...");

                        // Wait until Pulse is called on the workQueue object.
                        Monitor.Wait(workQueue);
                    }
                    else
                    {
                        // Obtain the next work item
                        workItem = workQueue.Dequeue();
                    }
                }
                finally
                {
                    // Always release the lock.
                    Monitor.Exit(workQueue);
                }

                // Process the work item if one was obtained.
                if (workItem != null)
                {
                    // Obtain a lock on the console and display a series 
                    // of messages.
                    lock (consoleGate)
                    {
                        for (int i = 0; i < 5; i++)
                        {
                            TraceMsg("Processing " + workItem);
                            Thread.Sleep(200);
                        }
                    }

                    // Reset the status of the local variable.
                    workItem = null;
                }
            }

            // This will only be reached if processWorkItems is false.
            TraceMsg("Terminating.");
        }

        public static void Main()
        {
            TraceMsg("Starting worker threads.");

            // Add an initial work item to the work queue.
            lock (workQueue)
            {
                workQueue.Enqueue("Work Item 1");
            }

            // Create and start three new worker threads running the 
            // ProcessWorkItems method.
            for (int count = 0; count < 3; count++)
            {
                (new Thread(ProcessWorkItems)).Start();
            }

            Thread.Sleep(1500);

            // The first time the user presses Enter, add a work item and 
            // active a single thread to process it.
            TraceMsg("Press Enter to pulse one waiting thread.");
            Console.ReadLine();

            // Acquire a lock on the workQueue object.
            lock (workQueue)
            {
                // Add a work item.
                workQueue.Enqueue("Work Item 2.");
                
                // Pulse 1 waiting thread.
                Monitor.Pulse(workQueue);
            }

            Thread.Sleep(2000);

            // The second time the user presses Enter, add three work items and 
            // activate three threads to process them.
            TraceMsg("Press Enter to pulse three waiting threads.");
            Console.ReadLine();

            // Acquire a lock on the workQueue object.
            lock (workQueue)
            {
                // Add work items to the work queue and activate worker threads.
                workQueue.Enqueue("Work Item 3.");
                Monitor.Pulse(workQueue);
                workQueue.Enqueue("Work Item 4.");
                Monitor.Pulse(workQueue);
                workQueue.Enqueue("Work Item 5.");
                Monitor.Pulse(workQueue);
            }

            Thread.Sleep(3500);

            // The third time the user presses Enter, signal the worker threads
            // to terminate and activate them all.
            TraceMsg("Press Enter to pulse all waiting threads.");
            Console.ReadLine();

            // Acquire a lock on the workQueue object.
            lock (workQueue)
            {
                // Signal that threads should terminate.
                processWorkItems = false;

                // Pulse all waiting threads.
                Monitor.PulseAll(workQueue);
            }

            Thread.Sleep(1000);

            // Wait to continue.
            TraceMsg("Main method complete. Press Enter.");
            Console.ReadLine();
        }
    }
}