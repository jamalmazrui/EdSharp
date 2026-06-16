using System;
using System.Threading;

namespace Apress.VisualCSharpRecipes.Chapter04
{
    class Recipe04_03
    {
        public static void Main() 
        {
            // Create the state object that is passed to the TimerHandler
            // method when it is triggered. In this case a message to display.
            string state = "Timer expired.";

            Console.WriteLine("{0} : Creating Timer.",
                DateTime.Now.ToString("HH:mm:ss.ffff"));

            // Create a Timer that fires first after 2 seconds and then every
            // second. Use an anonymouse method for the timer expiry handler.
            using (Timer timer =
                new Timer(delegate(object s) 
                            {Console.WriteLine("{0} : {1}",
                             DateTime.Now.ToString("HH:mm:ss.ffff"),s); 
                            }
                , state, 2000, 1000))
            {
                int period;

                // Read the new timer interval from the console until the
                // user enters 0 (zero). Invalid values use a default value
                // of 0, which will stop the example.
                do 
                {
                    try 
                    {
                        period = Int32.Parse(Console.ReadLine());
                    } 
                    catch (FormatException)
                    {
                        period = 0;
                    }

                    // Change the timer to fire using the new interval starting
                    // immediately.
                    if (period > 0) timer.Change(0, period);
                } while (period > 0);
            }

            // Wait to continue.
            Console.WriteLine("Main method complete. Press Enter.");
            Console.ReadLine();
        }
    }
}