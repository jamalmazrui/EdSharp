using System;
using System.Threading;
using System.Globalization;

namespace Apress.VisualCSharpRecipes.Chapter04
{
    class Recipe04_04
    {
        public static void Main(string[] args)
        {
            // Create a 30 second timespan
            TimeSpan waitTime = new TimeSpan(0, 0, 30);

            // Create a Timer that fires once at the specified time. Specify
            // an interval of -1 to stop the timer executing the method 
            // repeatedly. Use an anonymouse method for the timer expiry handler.
            new Timer(delegate(object s)
            {
                Console.WriteLine("Timer fired at {0}",
                DateTime.Now.ToString("HH:mm:ss.ffff"));
            }
                      , null, waitTime, new TimeSpan(-1));

            Console.WriteLine("Waiting for timer. Press Enter to terminate.");
            Console.ReadLine();
        }
    }
}