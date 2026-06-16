using System;

namespace Apress.VisualCSharpRecipes.Chapter02
{
    class Recipe02_08
    {
        public static void Main()
        {
            // Create a TimeSpan representing 2.5 days.
            TimeSpan timespan1 = new TimeSpan(2, 12, 0, 0);

            // Create a TimeSpan representing 4.5 days.
            TimeSpan timespan2 = new TimeSpan(4, 12, 0, 0);

            // create a TimeSpan representing 3.2 days 
            // using the static convenience method
            TimeSpan timespan3 = TimeSpan.FromDays(3.2);

            // Create a TimeSpan representing 1 week.
            TimeSpan oneWeek = timespan1 + timespan2;

            // Create a DateTime with the current date and time.
            DateTime now = DateTime.Now;

            // Create a DateTime representing 1 week ago.
            DateTime past = now - oneWeek;

            // Create a DateTime representing 1 week in the future.
            DateTime future = now + oneWeek;

            // Display the DateTime instances.
            Console.WriteLine("Now   : {0}", now);
            Console.WriteLine("Past  : {0}", past);
            Console.WriteLine("Future: {0}", future);

            // use the comparison operators
            Console.WriteLine("Now is greater than past: {0}", now > past);
            Console.WriteLine("Now is equal to future: {0}", now == future);

            // Wait to continue.
            Console.WriteLine("\nMain method complete. Press Enter");
            Console.ReadLine();
        }
    }
}
