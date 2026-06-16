using System;
using System.Threading;
using System.Threading.Tasks;

namespace Recipe15_04
{
    class Recipe15_04
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Press enter to start");
            Console.ReadLine();

            // define the data we want to process
            int[] numbersArray = { 100, 200, 300 };

            // configure the options
            ParallelOptions options = new ParallelOptions();
            options.MaxDegreeOfParallelism = 2;

            // process each data element in parallel
            Parallel.ForEach(numbersArray, options, baseNumber => printNumbers(baseNumber));

            Console.WriteLine("Tasks Completed, Press Enter");
            Console.ReadLine();
        }

        static void printNumbers(int baseNumber)
        {
            for (int i = baseNumber, j = baseNumber + 10; i < j; i++)
            {
                Console.WriteLine("Number: {0}", i);
                Thread.Sleep(100);
            }
        }
    }
}
