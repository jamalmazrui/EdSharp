using System;
using System.Threading;

namespace Apress.VisualCSharpRecipes.Chapter04
{
    class Recipe04_05
    {
        // A method that is executed when the AutoResetEvent is signaled
        // or the wait operation times out.
        private static void EventHandler(object state, bool timedout) 
        {
            // Display appropriate message to the console based on whether 
            // the wait timed out or the AutoResetEvent was signaled.
            if (timedout) 
            {
                Console.WriteLine("{0} : Wait timed out.",
                    DateTime.Now.ToString("HH:mm:ss.ffff"));
            } 
            else 
            {
                Console.WriteLine("{0} : {1}", 
                    DateTime.Now.ToString("HH:mm:ss.ffff"), state);
            }
        }

        public static void Main() 
        {
            // Create the new AutoResetEvent in an unsignaled state. 
            AutoResetEvent autoEvent = new AutoResetEvent(false);

            // Create the state object that is passed to the event handler
            // method when it is triggered. In this case a message to display.
            string state = "AutoResetEvent signaled.";

            // Register the EventHandler method to wait for the AutoResetEvent to
            // be signaled. Set a time-out of 10 seconds, and configure the wait 
            // operation to reset after activation (last argument).
            RegisteredWaitHandle handle = ThreadPool.RegisterWaitForSingleObject(
                autoEvent, EventHandler, state, 10000, false);

            Console.WriteLine("Press ENTER to signal the AutoResetEvent" +
                " or enter \"Cancel\" to unregister the wait operation.");

            while (Console.ReadLine().ToUpper() != "CANCEL") 
            {
                // If "Cancel" has not been entered into the console, signal
                // the AutoResetEvent, which will cause the EventHandler
                // method to execute. The AutoResetEvent will automatically
                // revert to an unsignaled state.
                autoEvent.Set();
            }

            // Unregister the wait operation.
            Console.WriteLine("Unregistering wait operation.");
            handle.Unregister(null);

            // Wait to continue.
            Console.WriteLine("Main method complete. Press Enter.");
            Console.ReadLine();
        }
    }
}