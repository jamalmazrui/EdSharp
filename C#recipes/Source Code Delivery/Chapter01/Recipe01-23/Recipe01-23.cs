using System;

namespace Apress.VisualCSharpRecipes.Chapter01
{
    public class Recipe01_23
    {
        public static EventHandler MyEvent;

        static void Main(string[] args)
        {
            // use a named method to register for the event
            MyEvent += new EventHandler(EventHandlerMethod);

            // use an anonymous delegate to register for the event
            MyEvent += new EventHandler(delegate(object sender, EventArgs eventargs)
            {
                Console.WriteLine("Anonymous delegate called");
            });

            // use a lamda expression to register for the event
            MyEvent += new EventHandler((sender, eventargs) =>
            {
                Console.WriteLine("Lamda expression called");
            });

            Console.WriteLine("Raising the event");
            MyEvent.Invoke(new object(), new EventArgs());

            Console.WriteLine("\nMain method complete. Press Enter.");
            Console.ReadLine();
        }

        static void EventHandlerMethod(object sender, EventArgs args)
        {
            Console.WriteLine("Named method called");
        }
    }
}
