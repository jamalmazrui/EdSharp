using System;

namespace Apress.VisualCSharpRecipes.Chapter01
{
    class HelloWorld
    {
        public static void Main()
        {
            ConsoleUtils.WriteString("Hello, world");

            Console.WriteLine("\nMain method complete. Press Enter.");
            Console.ReadLine();
        }
    }
}

