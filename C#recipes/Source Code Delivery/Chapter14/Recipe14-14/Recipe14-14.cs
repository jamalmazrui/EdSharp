using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Diagnostics;
using System.Security.Principal;

namespace Recipe14_14
{
    class Recipe14_14
    {
        static void Main(string[] args)
        {

            if (!checkElevatedPrivilege())
            {
                Console.WriteLine("This recipe requires administrator rights");
                Console.ReadLine();
                Environment.Exit(1);
            }

            // define the category name for the performance counters
            string categoryName = "Recipe 14-13 Performance Counters";

            // open the counters for reading
            PerformanceCounter perfCounter1 = new PerformanceCounter();
            perfCounter1.CategoryName = categoryName;
            perfCounter1.CounterName = "Number of Items Counter";           

            PerformanceCounter perfCounter2 = new PerformanceCounter();
            perfCounter2.CategoryName = categoryName;
            perfCounter2.CounterName = "Average Timer Counter";
   
            while (true)
            {
                float value1 = perfCounter1.NextValue();
                Console.WriteLine("Value for first counter: {0}", value1);
                float value2 = perfCounter2.NextValue();
                Console.WriteLine("Value for second counter: {0}", value2);

                // put the thread to sleep for a second
                System.Threading.Thread.Sleep(1000);
            }
        }

        static bool checkElevatedPrivilege()
        {
            WindowsIdentity winIdentity = WindowsIdentity.GetCurrent();
            WindowsPrincipal winPrincipal = new WindowsPrincipal(winIdentity);
            return winPrincipal.IsInRole(WindowsBuiltInRole.Administrator);
        }
    }
}
