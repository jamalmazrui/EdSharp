using System;
using System.Threading;
using System.Threading.Tasks;

namespace Recipe15_01
{
    class Recipe15_01
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Press enter to start");
            Console.ReadLine();

            // invoke the methods we want to run
            Parallel.Invoke(
                new Action(writeDays),
                new Action(writeMonths),
                new Action(writeCities)
            );

            // Wait to continue.
            Console.WriteLine("\nMain method complete. Press Enter");
            Console.ReadLine();
        }

       static void writeDays()
        {
            string[] daysArray = { "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", 
                                     "Saturday", "Sunday" };
            foreach (string day in daysArray)
            {
                Console.WriteLine("Day of the Week: {0}", day);
                Thread.Sleep(500);
            }
        }

        static void writeMonths()
        {
            string[] monthsArray = { "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul",
                                       "Aug", "Sep", "Oct", "Nov", "Dec" };
            foreach (string month in monthsArray)
            {
                Console.WriteLine("Month: {0}", month);
                Thread.Sleep(500);
            }
        }

        static void writeCities()
        {
            string[] citiesArray = { "London", "New York", "Paris", "Tokyo", "Sydney" };
            foreach (string city in citiesArray)
            {
                Console.WriteLine("City: {0}", city);
                Thread.Sleep(500);
            }

        }
    }
}
