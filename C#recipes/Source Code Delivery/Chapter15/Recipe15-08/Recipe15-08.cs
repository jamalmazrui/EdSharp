using System;
using System.Threading;
using System.Threading.Tasks;

namespace Recipe15_08
{
    class Recipe15_08
    {
        static void Main(string[] args)
        {
            // create the token source
            CancellationTokenSource tokenSource = new CancellationTokenSource();
            // create the cancellation token
            CancellationToken token = tokenSource.Token;

            // create the task
            Task task = Task.Factory.StartNew(() => printNumbers(token), token);
            // register the task with the token
            token.Register(() => notifyTaskCancelled());

            // wait for the user to request cancellation
            Console.WriteLine("Press enter to cancel token");
            Console.ReadLine();

            // cancelling
            tokenSource.Cancel();
        }

        static void notifyTaskCancelled()
        {
            Console.WriteLine("Task cancellation requested");
        }

        static void printNumbers(CancellationToken token)
        {
            int i = 0;
            while (!token.IsCancellationRequested)
            {
                Console.WriteLine("Number {0}", i++);
                Thread.Sleep(500);
            }
            throw new OperationCanceledException(token);
        }
    }
}
