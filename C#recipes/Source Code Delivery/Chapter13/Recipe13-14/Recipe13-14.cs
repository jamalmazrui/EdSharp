using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Apress.VisualCSharpRecipes.Chapter13
{
    class Recipe13_14
    {
        static void Main(string[] args)
        {
            printMessage();
            printMessage("Allen");
            printMessage("Bob", "Goodbye");
            printMessage("Joe", "Help", true);

            // Wait to continue.
            Console.WriteLine(Environment.NewLine);
            Console.WriteLine("Main method complete. Press Enter");
            Console.ReadLine();
        }

        static void printMessage(string from = "Adam", string message = "Hello", bool urgent = false)
        {
            Console.WriteLine("From: {0}, Urgent: {1}, Message: {2}", from, message, urgent);
        }
    }
}
