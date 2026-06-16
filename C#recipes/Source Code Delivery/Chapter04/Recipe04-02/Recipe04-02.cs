using System;
using System.Threading;
using System.Collections;

namespace Apress.VisualCSharpRecipes.Chapter04
{
    class Recipe04_02
    {
        // A utility method for displaying useful trace information to the 
        // console along with details of the current thread.
        private static void TraceMsg(DateTime time, string msg)
        {
            Console.WriteLine("[{0,3}/{1}] - {2} : {3}",
                Thread.CurrentThread.ManagedThreadId,
                Thread.CurrentThread.IsThreadPoolThread ? "pool" : "fore",
                time.ToString("HH:mm:ss.ffff"), msg);
        }

        // A delegate that allows you to perform asynchronous execution of 
        // LongRunningMethod.
        public delegate DateTime AsyncExampleDelegate(int delay, string name);

        // A simulated long running method.
        public static DateTime LongRunningMethod(int delay, string name)
        {
            TraceMsg(DateTime.Now, name + " example - thread starting.");

            // Simulate time consuming processing.
            Thread.Sleep(delay);

            TraceMsg(DateTime.Now, name + " example - thread stopping.");

            // Return the method's completion time.
            return DateTime.Now;
        }

        // This method executes LongRunningMethod asynchronously and continues
        // with other processing. Once the processing is complete, the method
        // blocks until LongRunningMethod completes.
        public static void BlockingExample()
        {
            Console.WriteLine(Environment.NewLine +
                "*** Running Blocking Example ***");

            // Invoke LongRunningMethod asynchronously. Pass null for both the 
            // callback delegate and the asynchronous state object.
            AsyncExampleDelegate longRunningMethod = LongRunningMethod;

            IAsyncResult asyncResult = longRunningMethod.BeginInvoke(2000,
                "Blocking", null, null);

            // Perform other processing until ready to block.
            for (int count = 0; count < 3; count++)
            {
                TraceMsg(DateTime.Now, 
                    "Continue processing until ready to block...");

                Thread.Sleep(200);
            }

            // Block until the asynchronous method completes. 
            TraceMsg(DateTime.Now,
                "Blocking until method is complete...");

            // Obtain the completion data for the asynchronous method.
            DateTime completion = DateTime.MinValue;

            try
            {
                completion = longRunningMethod.EndInvoke(asyncResult);
            }
            catch
            {
                // Catch and handle those exceptions you would if calling
                // LongRunningMethod directly.
            }

            // Display completion information
            TraceMsg(completion,"Blocking example complete.");
        }

        // This method executes LongRunningMethod asynchronously and then 
        // enters a polling loop until LongRunningMethod completes.
        public static void PollingExample()
        {
            Console.WriteLine(Environment.NewLine +
                "*** Running Polling Example ***");

            // Invoke LongRunningMethod asynchronously. Pass null for both the 
            // callback delegate and the asynchronous state object.
            AsyncExampleDelegate longRunningMethod = LongRunningMethod;

            IAsyncResult asyncResult = longRunningMethod.BeginInvoke(2000,
                "Polling", null, null);

            // Poll the asynchronous method to test for completion. If not 
            // complete sleep for 300ms before polling again.
            TraceMsg(DateTime.Now, "Poll repeatedly until method is complete.");

            while (!asyncResult.IsCompleted)
            {
                TraceMsg(DateTime.Now, "Polling...");
                Thread.Sleep(300);
            }

            // Obtain the completion data for the asynchronous method.
            DateTime completion = DateTime.MinValue;

            try
            {
                completion = longRunningMethod.EndInvoke(asyncResult);
            }
            catch
            {
                // Catch and handle those exceptions you would if calling
                // LongRunningMethod directly.
            }

            // Display completion information
            TraceMsg(completion, "Polling example complete.");
        }

        // This method executes LongRunningMethod asynchronously and then 
        // uses a WaitHandle to wait efficiently until LongRunningMethod 
        // completes. Use of a timeout allows the method to break out of
        // waiting in order to update the user interface or fail if the 
        // asynchronous method is taking too long.
        public static void WaitingExample()
        {
            Console.WriteLine(Environment.NewLine +
                "*** Running Waiting Example ***");

            // Invoke LongRunningMethod asynchronously. Pass null for both the 
            // callback delegate and the asynchronous state object.
            AsyncExampleDelegate longRunningMethod = LongRunningMethod;

            IAsyncResult asyncResult = longRunningMethod.BeginInvoke(2000,
                "Waiting", null, null);

            // Wait for the asynchronous method to complete. Time out after 
            // 300ms and display status to the console before continuing to
            // wait.
            TraceMsg(DateTime.Now, "Waiting until method is complete...");

            while (!asyncResult.AsyncWaitHandle.WaitOne(300, false))
            {
                TraceMsg(DateTime.Now, "Wait timeout...");
            }

            // Obtain the completion data for the asynchronous method.
            DateTime completion = DateTime.MinValue;

            try
            {
                completion = longRunningMethod.EndInvoke(asyncResult);
            }
            catch
            {
                // Catch and handle those exceptions you would if calling
                // LongRunningMethod directly.
            }

            // Display completion information
            TraceMsg(completion, "Waiting example complete.");
        }

        // This method executes LongRunningMethod asynchronously multiple 
        // times and then uses an array of WaitHandle objects to wait 
        // efficiently until all of the methods are complete. Use of
        // a timeout allows the method to break out of waiting in order
        // to update the user interface or fail if the asynchronous 
        // method is taking too long.
        public static void WaitAllExample()
        {
            Console.WriteLine(Environment.NewLine +
                "*** Running WaitAll Example ***");

            // An ArrayList to hold the IAsyncResult instances for each of the 
            // asynchronous methods started.
            ArrayList asyncResults = new ArrayList(3);

            // Invoke three LongRunningMethods asynchronously. Pass null for 
            // both the callback delegate and the asynchronous state object.
            // Add the IAsyncResult instance for each method to the ArrayList.
            AsyncExampleDelegate longRunningMethod = LongRunningMethod;

            asyncResults.Add(longRunningMethod.BeginInvoke(3000,
                "WaitAll 1", null, null));

            asyncResults.Add(longRunningMethod.BeginInvoke(2500,
                "WaitAll 2", null, null));

            asyncResults.Add(longRunningMethod.BeginInvoke(1500,
                "WaitAll 3", null, null));

            // Create an array of WaitHandle objects that will be used to wait 
            // for the completion of all of the asynchronous methods.
            WaitHandle[] waitHandles = new WaitHandle[3];

            for (int count = 0; count < 3; count++)
            {
                waitHandles[count] =
                    ((IAsyncResult)asyncResults[count]).AsyncWaitHandle;
            }

            // Wait for all three asynchronous method to complete. Time out
            // after 300ms and display status to the console before continuing
            // to wait.
            TraceMsg(DateTime.Now, "Waiting until all 3 methods are complete...");

            while (!WaitHandle.WaitAll(waitHandles, 300, false))
            {
                TraceMsg(DateTime.Now, "WaitAll timeout...");
            }

            // Inspect the completion data for each method and determine the
            // time at which the final method completed.
            DateTime completion = DateTime.MinValue;

            foreach (IAsyncResult result in asyncResults)
            {
                try
                {
                    DateTime time = longRunningMethod.EndInvoke(result);
                    if (time > completion) completion = time;
                }
                catch
                {
                    // Catch and handle those exceptions you would if calling
                    // LongRunningMethod directly.
                }
            }

            // Display completion information
            TraceMsg(completion, "WaitAll example complete.");
        }

        // This method executes LongRunningMethod asynchronously and passes 
        // an AsyncCallback delegate instance. The referenced CallbackHandler
        // method is called automatically when the asynchronous method 
        // completes leaving this method free to continue processing.
        public static void CallbackExample()
        {
            Console.WriteLine(Environment.NewLine +
                "*** Running Callback Example ***");

            // Invoke LongRunningMethod asynchronously. Pass an AsyncCallback
            // delegate instance referencing the CallbackHandler method which
            // will be called automatically when the asynchronous method 
            // completes. Pass a reference to the AsyncExampleDelegate delegate 
            // instance as asynchronous state; otherwise, the callback method
            // has no access to the delegate instance in order to call 
            // EndInvoke.
            AsyncExampleDelegate longRunningMethod = LongRunningMethod;

            IAsyncResult asyncResult = longRunningMethod.BeginInvoke(2000,
                "Callback", CallbackHandler, longRunningMethod);

            // Continue with other processing.
            for (int count = 0; count < 15; count++)
            {
                TraceMsg(DateTime.Now, "Continue processing...");
                Thread.Sleep(200);
            }
        }

        // A method to handle asynchronous completion using callbacks.
        public static void CallbackHandler(IAsyncResult result)
        {
            // Extract the reference to the AsyncExampleDelegate instance 
            // from the IAsyncResult instance. This allows us to obtain the 
            // completion data.
            AsyncExampleDelegate longRunningMethod =
                (AsyncExampleDelegate)result.AsyncState;

            // Obtain the completion data for the asynchronous method.
            DateTime completion = DateTime.MinValue;

            try
            {
                completion = longRunningMethod.EndInvoke(result);
            }
            catch
            {
                // Catch and handle those exceptions you would if calling
                // LongRunningMethod directly.
            }

            // Display completion information
            TraceMsg(completion, "Callback example complete.");
        }

        public static void Main()
        {
            // Demonstrate the various approaches to asynchronous method completion.
            BlockingExample();
            PollingExample();
            WaitingExample();
            WaitAllExample();
            CallbackExample();

            // Wait to continue.
            Console.WriteLine(Environment.NewLine);
            Console.WriteLine("Main method complete. Press Enter.");
            Console.ReadLine();
        }
    }
}
