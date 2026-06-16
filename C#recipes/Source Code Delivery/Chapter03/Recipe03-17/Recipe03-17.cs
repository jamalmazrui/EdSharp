using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Reflection;

namespace Apress.VisualCSharpRecipes.Chapter03
{
    class Recipe03_17
    {
        static void Main(string[] args)
        {
            // create an instance of this type
            dynamic myInstance = new Recipe03_17();

            // call the method dynamically
            myInstance.printMessage("hello", 37, 'c');

            // Wait to continue.
            Console.WriteLine("\nMain method complete. Press Enter.");
            Console.ReadLine();
        }

        public void printMessage(string param1, int param2, char param3)
        {
            Console.WriteLine("PrintMessage {0} {1} {2}", param1, param2, param3);
        }
    }
}
