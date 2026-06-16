using System;
using System.Collections.Concurrent;
using System.Threading.Tasks;

namespace Recipe15_09
{
    class Recipe15_09
    {
        static void Main(string[] args)
        {
            // create a concurrent collection
            ConcurrentStack<int> cStack = new ConcurrentStack<int>();

            // create tasks that will use the stack
            Task task1 = Task.Factory.StartNew(() => addNumbersToCollection(cStack));
            Task task2 = Task.Factory.StartNew(() => addNumbersToCollection(cStack));
            Task task3 = Task.Factory.StartNew(() => addNumbersToCollection(cStack));

            // wait for all of the tasks to complete
            Task.WaitAll(task1, task2, task3);

            // report how many items there are in the stack
            Console.WriteLine("There are {0} items in the collection", cStack.Count);

            // Wait to continue.
            Console.WriteLine("\nMain method complete. Press Enter");
            Console.ReadLine();
        }

        static void addNumbersToCollection(ConcurrentStack<int> stack)
        {
            for (int i = 0; i < 1000; i++)
            {
                stack.Push(i);
            }
        }
    }
}
