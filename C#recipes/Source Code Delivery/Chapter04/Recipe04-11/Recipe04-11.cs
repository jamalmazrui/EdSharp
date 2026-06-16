using System;
using System.Threading;

namespace Apress.VisualCSharpRecipes.Chapter04
{
    class Recipe04_11
    {
        public static void Main()
        {
            int firstInt = 2500;
            int secondInt = 8000;

            Console.WriteLine("firstInt initial value = {0}", firstInt);
            Console.WriteLine("secondInt initial value = {0}", secondInt);

            // Decrement firstInt in a thread safe manner.
            // The thread-safe equivalent of firstInt = firstInt + 1.
            Interlocked.Decrement(ref firstInt);

            Console.WriteLine(Environment.NewLine); 
            Console.WriteLine("firstInt after decrement = {0}", firstInt);

            // Increment secondInt in a thread safe manner.
            // The thread-safe equivalent of secondInt = secondInt + 1.
            Interlocked.Increment(ref secondInt);

            Console.WriteLine("secondInt after increment = {0}", secondInt);

            // Add the firstInt and secondInt values and store the result in 
            // firstInt.
            // The thread-safe equivalent of firstInt = firstInt + secondInt.
            Interlocked.Add(ref firstInt, secondInt);

            Console.WriteLine(Environment.NewLine); 
            Console.WriteLine("firstInt after Add = {0}", firstInt);
            Console.WriteLine("secondInt after Add = {0}", secondInt);

            // Exchange the value of firstInt with secondInt.
            // The thread-safe equivalenet of secondInt = firstInt.
            Interlocked.Exchange(ref secondInt, firstInt);

            Console.WriteLine(Environment.NewLine); 
            Console.WriteLine("firstInt after Exchange = {0}", firstInt);
            Console.WriteLine("secondInt after Exchange = {0}", secondInt);

            // Compare firstInt with secondInt and if they are equal set 
            // firstInt to 5000.
            // The thread-safe equivalenet of:
            //     if (firstInt == secondInt) firstInt = 5000
            Interlocked.CompareExchange(ref firstInt, 5000, secondInt);

            Console.WriteLine(Environment.NewLine); 
            Console.WriteLine("firstInt after CompareExchange = {0}", firstInt);
            Console.WriteLine("secondInt after CompareExchange = {0}", secondInt);

            // Wait to continue.
            Console.WriteLine(Environment.NewLine);
            Console.WriteLine("Main method complete. Press Enter.");
            Console.ReadLine();
        }
    }
}
