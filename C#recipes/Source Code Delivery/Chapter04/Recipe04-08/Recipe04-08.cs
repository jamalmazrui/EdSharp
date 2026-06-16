using System;
using System.Threading;

namespace Apress.VisualCSharpRecipes.Chapter04
{
    class Recipe04_08
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
        // The method waits on the EventWaitHandle before displaying a message
        // to the console, then waits 2 seconds and loops.
        private static void DisplayMessage()
        {
            // Obtain a handle to the EventWaitHandle with the name "EventExample".
            EventWaitHandle eventHandle = 
                EventWaitHandle.OpenExisting("EventExample");

            TraceMsg("DisplayMessage Started.");

            while (!terminate)
            {
                // Wait on the EventWaitHandle, timeout after 2 seconds. WaitOne 
                // returns true if the event is signalled; otherwise false. The 
                // first time through, the message will be displayed immediately
                // because the EventWaitHandle was created in a signalled state.
                if (eventHandle.WaitOne(2000, true))
                {
                    TraceMsg("EventWaitHandle In Signaled State.");
                }
                else
                {
                    TraceMsg("WaitOne Timed Out -- " + 
                        "EventWaitHandle In Unsignaled State.");
                }
                Thread.Sleep(2000);
            }

            TraceMsg("Thread Terminating.");
        }

        public static void Main()
        {
            // Create a new EventWaitHandle with an initial signaled state, in 
            // manual mode, with the name "EventExample".
            using (EventWaitHandle eventWaitHandle =
                new EventWaitHandle(true, EventResetMode.ManualReset, 
                "EventExample"))
            {
                // Create and start a new thread running the DisplayMesssage 
                // method.
                TraceMsg("Starting DisplayMessageThread.");
                Thread trd = new Thread(DisplayMessage);
                trd.Start();

                // Allow the EventWaitHandle to be toggled between a signalled and
                // unsignalled state up to 3 time before ending.
                for (int count = 0; count < 3; count++)
                {
                    // Wait for Enter to be pressed.
                    Console.ReadLine();

                    // We need to toggle the event. The only way to know the 
                    // current state is to wait on it with a 0 (zero) timeout 
                    // and test the result.
                    if (eventWaitHandle.WaitOne(0, true))
                    {
                        TraceMsg("Switching Event To UnSignaled State.");

                        // Event is signalled, so unsignal it.
                        eventWaitHandle.Reset();
                    }
                    else
                    {
                        TraceMsg("Switching Event To Signaled State.");

                        // Event is unsignalled, so signal it.
                        eventWaitHandle.Set();
                    }
                }

                // Terminate the DisplayMessage thread and wait for it to 
                // complete before disposing of the EventWaitHandle.
                terminate = true;
                eventWaitHandle.Set();
                trd.Join(5000);
            }

            // Wait to continue.
            Console.WriteLine(Environment.NewLine);
            Console.WriteLine("Main method complete. Press Enter.");
            Console.ReadLine();
        }
    }
}
