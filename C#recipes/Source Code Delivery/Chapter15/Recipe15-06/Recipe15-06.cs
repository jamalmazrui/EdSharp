using System;
using System.Threading;
using System.Threading.Tasks;

namespace Recipe15_06
{
    class Recipe15_06
    {
        static void Main(string[] args)
        {
            // create the barrier
            Barrier myBarrier = new Barrier(3, (barrier) => notifyPhaseEnd(barrier));
  
            Task task1 = Task.Factory.StartNew(() => cooperatingAlgorithm(1, myBarrier));
            Task task2 = Task.Factory.StartNew(() => cooperatingAlgorithm(2, myBarrier));
            Task task3 = Task.Factory.StartNew(() => cooperatingAlgorithm(3, myBarrier));

            // wait for all of the tasks to complete
            Task.WaitAll(task1, task2, task3);

            // Wait to continue.
            Console.WriteLine("\nMain method complete. Press Enter");
            Console.ReadLine();
        }

        static void cooperatingAlgorithm(int taskid, Barrier barrier)
        {
            Console.WriteLine("Running algorithm for task {0}", taskid);

            // perform phase one and wait at the barrier
            performPhase1(taskid);
            barrier.SignalAndWait();

            // perform phase two and wait at the barrier
            performPhase2(taskid);
            barrier.SignalAndWait();
        }

        static void performPhase1(int taskid)
        {
            Console.WriteLine("Phase one performed for task {0}", taskid);
        }

        static void performPhase2(int taskid)
        {
            Console.WriteLine("Phase two performed for task {0}", taskid);
        }

        static void notifyPhaseEnd(Barrier barrier)
        {
            Console.WriteLine("Phase has concluded");
        }
    }
}
