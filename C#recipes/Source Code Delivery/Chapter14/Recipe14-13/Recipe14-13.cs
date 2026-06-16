using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Security.Principal;

namespace Recipe14_13
{
    class Recipe14_13
    {

        [DllImport("Kernel32.dll")]
        public static extern void QueryPerformanceCounter(ref long ticks);

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

            if (!PerformanceCounterCategory.Exists(categoryName))
            {
                Console.WriteLine("Creating counters.");

                // we need to create the category
                CounterCreationDataCollection counterCollection
                    = new CounterCreationDataCollection();

                // create the individual counters
                CounterCreationData counter1 = new CounterCreationData();
                counter1.CounterType = PerformanceCounterType.NumberOfItems32;
                counter1.CounterName = "Number of Items Counter";
                counter1.CounterHelp = "A sample 32-bit number counter";

                CounterCreationData counter2 = new CounterCreationData();
                counter2.CounterType = PerformanceCounterType.AverageTimer32;
                counter2.CounterName = "Average Timer Counter";
                counter2.CounterHelp = "A sample average timer counter";

                CounterCreationData counter3 = new CounterCreationData();
                counter3.CounterType = PerformanceCounterType.AverageBase;
                counter3.CounterName = "Average Base Counter";
                counter3.CounterHelp = "A sample average base counter";

                // add the counters to the collection
                counterCollection.Add(counter1);
                counterCollection.Add(counter2);
                counterCollection.Add(counter3);

                // create the counters category
                PerformanceCounterCategory.Create(categoryName,
                    "Category for Visual C# Recipe 14-13",
                    PerformanceCounterCategoryType.SingleInstance,
                    counterCollection);
            }
            else
            {
                Console.WriteLine("Counters already exist.");
            }

            // open the counters for reading
            PerformanceCounter perfCounter1 = new PerformanceCounter();
            perfCounter1.CategoryName = categoryName;
            perfCounter1.CounterName = "Number of Items Counter";
            perfCounter1.ReadOnly = false;           

            PerformanceCounter perfCounter2 = new PerformanceCounter();
            perfCounter2.CategoryName = categoryName;
            perfCounter2.CounterName = "Average Timer Counter";
            perfCounter2.ReadOnly = false;

            PerformanceCounter perfCounter3 = new PerformanceCounter();
            perfCounter3.CategoryName = categoryName;
            perfCounter3.CounterName = "Average Base Counter";
            perfCounter3.ReadOnly = false;

            // create a number generator to produce values
            Random numberGenerator = new Random();

            // enter a loop to update the values every second
            long startTickCount = 0, endTickCount = 0;
            while (true)
            {
                // get the high frequency tick count
                QueryPerformanceCounter(ref startTickCount);
                // put the thread to sleep for up to a second
                System.Threading.Thread.Sleep(numberGenerator.Next(1000));
                // get the high frequency tick count again
                QueryPerformanceCounter(ref endTickCount);

                Console.WriteLine("Updating counter values.");
                perfCounter1.Increment();
                perfCounter2.IncrementBy(endTickCount - startTickCount);
                perfCounter3.Increment();
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
