using System;
using System.Threading;
using System.Threading.Tasks;

namespace Recipe15_02
{
    class Recipe15_02
    {
        static void Main(string[] args)
        {

            Console.WriteLine("Press enter to start");
            Console.ReadLine();

            // create the tasks
            Task<int> task1 = Task<int>.Factory.StartNew(() => writeDays());
            Task<int> task2 = Task<int>.Factory.StartNew(() => writeMonths());
            Task<int> task3 = Task<int>.Factory.StartNew(() => writeCities());

            // get the results and write them out
            Console.WriteLine("{0} days were written", task1.Result);
            Console.WriteLine("{0} months were written", task2.Result);
            Console.WriteLine("{0} cities were written", task3.Result);

            // Wait to continue.
            Console.WriteLine("\nMain method complete. Press Enter");
            Console.ReadLine();
        }

       static int writeDays()
        {
            string[] daysArray = { "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", 
                                     "Saturday", "Sunday" };
            foreach (string day in daysArray)
            {
                Console.WriteLine("Day of the Week: {0}", day);
                Thread.Sleep(500);
            }
            return daysArray.Length;
        }

        static int writeMonths()
        {
            string[] monthsArray = { "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul",
                                       "Aug", "Sep", "Oct", "Nov", "Dec" };
            foreach (string month in monthsArray)
            {
                Console.WriteLine("Month: {0}", month);
                Thread.Sleep(500);
            }
            return monthsArray.Length;
        }

        static int writeCities()
        {
            string[] citiesArray = { "London", "New York", "Paris", "Tokyo", "Sydney" };
            foreach (string city in citiesArray)
            {
                Console.WriteLine("City: {0}", city);
                Thread.Sleep(500);
            }
            return citiesArray.Length;
        }
    }
}
