using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Collections.Concurrent;
using System.Threading.Tasks;

namespace Apress.VisualCSharpRecipes.Chapter13
{
    class Recipe13_12
    {
        static void Main(string[] args)
        {
            // create the collection
            BlockingCollection<string> bCollection
                = new BlockingCollection<string>(new ConcurrentQueue<string>(), 3);

            // start the consumer
            Task consumerTask = Task.Factory.StartNew(
                () => performConsumerTasks(bCollection));
            // start the producers
            Task producer0 = Task.Factory.StartNew(
                () => performProducerTasks(bCollection, 0));
            Task producer1 = Task.Factory.StartNew(
                () => performProducerTasks(bCollection, 1));
            Task producer2 = Task.Factory.StartNew(
                () => performProducerTasks(bCollection, 2));
            Task producer3 = Task.Factory.StartNew(
                () => performProducerTasks(bCollection, 3));

            // wait for the producers to complete
            Task.WaitAll(producer0, producer1, producer2, producer3);
            Console.WriteLine("Producer tasks complete.");

            // indicate that we will not add anything further
            bCollection.CompleteAdding();

            // wait for the consumer task to complete
            consumerTask.Wait();

            // Wait to continue.
            Console.WriteLine(Environment.NewLine);
            Console.WriteLine("Main method complete. Press Enter");
            Console.ReadLine();
        }

        static void performConsumerTasks(BlockingCollection<string> collection)
        {
            Console.WriteLine("Consumer started");
            foreach (string collData in collection.GetConsumingEnumerable())
            {
                // write out the data
                Console.WriteLine("Consumer got message {0}", collData);
            }
            Console.WriteLine("Consumer completed");
        }

        static void performProducerTasks(BlockingCollection<string> collection, int taskID)
        {
            Console.WriteLine("Producer started");
            for (int i = 0; i < 100; i++)
            {
                // put something into the collection
                collection.Add("TaskID " + taskID + " Message" + i++);
                Console.WriteLine("Task {0} wrote message {1}", taskID, i);
            }
        }
    }
}
