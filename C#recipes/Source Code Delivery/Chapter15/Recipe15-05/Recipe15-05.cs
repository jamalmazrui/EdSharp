using System;
using System.Threading.Tasks;

namespace Recipe15_05
{
    class Recipe15_05
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Press enter to start");
            Console.ReadLine();

            // create the set of tasks
            Task<int> firstTask = new Task<int>(() => sumAndPrintNumbers(100));
            Task secondTask = firstTask.ContinueWith(parent => printTotal(parent));
            Task thirdTask = secondTask.ContinueWith(parent => printMessage());

            // start the first task
            firstTask.Start();
            
            // read a line to keep the process alive
            Console.WriteLine("Press enter to finish");
            Console.ReadLine();
        }

        static int sumAndPrintNumbers(int baseNumber)
        {
            Console.WriteLine("sum&print called for {0}", baseNumber);
            int total = 0;
            for (int i = baseNumber, j = baseNumber + 10; i < j; i++)
            {
                Console.WriteLine("Number: {0}", i);
                total += i;
            }
            return total;
        }

        static void printTotal(Task<int> parentTask)
        {
            Console.WriteLine("Total is {0}", parentTask.Result);
        }

        static void printMessage()
        {
            Console.WriteLine("Message from third task");
        }
    }
}
