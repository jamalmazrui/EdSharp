using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Recipe15_07
{
    class Recipe15_07
    {
        static void Main(string[] args)
        {
            // create two tasks, one with a null param
            Task goodTask = Task.Factory.StartNew(() => performTask("good"));
            Task badTask = Task.Factory.StartNew(() => performTask("bad"));

            try
            {
                Task.WaitAll(goodTask, badTask);
            }
            catch (AggregateException aggex)
            {
                aggex.Handle(ex => handleException(ex));
            }

            // Wait to continue.
            Console.WriteLine("\nMain method complete. Press Enter");
            Console.ReadLine();
        }

        static bool handleException(Exception exception)
        {
            Console.WriteLine("Processed Exception");
            Console.WriteLine(exception);
            // return true to indicate we have handled the exception
            return true;
        }

        static void performTask(string label)
        {
            if (label == "bad")
            {
                Console.WriteLine("About to throw exception.");
                throw new ArgumentOutOfRangeException("label");
            }
            else
            {
                Console.WriteLine("performTask for label: {0}", label);
            }
        }
    }
}
