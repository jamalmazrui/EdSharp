using System;
using System.IO;
using System.Threading;

namespace Apress.VisualCSharpRecipes.Chapter05
{
    static class Recipe05_09
    {
        static void Main(string[] args)
        {
            // Create a test file.
            using (FileStream fs = new FileStream("test.txt", FileMode.Create))
            {
                fs.SetLength(100000);
            }

            // Start the asynchronous file processor on another thread.
            AsyncProcessor asyncIO = new AsyncProcessor("test.txt");
            asyncIO.StartProcess();

            // At the same time, do some other work.
            // In this example, we simply loop for 10 seconds.
            DateTime startTime = DateTime.Now;
            while (DateTime.Now.Subtract(startTime).TotalSeconds < 2)
            {
                Console.WriteLine("[MAIN THREAD]: Doing some work.");

                // Pause to simulate a time-consuming operation.
                Thread.Sleep(TimeSpan.FromMilliseconds(100));
            }

            Console.WriteLine("[MAIN THREAD]: Complete.");
            Console.ReadLine();

            // Remove the test file.
            File.Delete("test.txt");
        }
    }
}
