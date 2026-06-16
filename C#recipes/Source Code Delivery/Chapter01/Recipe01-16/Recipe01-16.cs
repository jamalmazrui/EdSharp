using System;

namespace Apress.VisualCSharpRecipes.Chapter01
{
    public class Recipe01_16
    {
        static void Main(string[] args)
        {
            // Display the standard Console.
            Console.Title = "Standard Console";
            Console.WriteLine("Press Enter to change the Console's appearance.");
            Console.ReadLine();

            // Change the Console appearance and redisplay.
            Console.Title = "Colored Text"; 
            Console.ForegroundColor = ConsoleColor.Red;
            Console.BackgroundColor = ConsoleColor.Green;
            Console.WriteLine("Press Enter to change the Console's appearance.");
            Console.ReadLine();

            // Change the Console appearance and redisplay.
            Console.Title = "Cleared / Colored Console";
            Console.ForegroundColor = ConsoleColor.Blue;
            Console.BackgroundColor = ConsoleColor.Yellow;
            Console.Clear();
            Console.WriteLine("Press Enter to change the Console's appearance.");
            Console.ReadLine();

            // Change the Console appearance and redisplay.
            Console.Title = "Resized Console";
            Console.ResetColor();
            Console.Clear();
            Console.SetWindowSize(100, 50);
            Console.BufferHeight = 500;
            Console.BufferWidth = 100;
            Console.CursorLeft = 20;
            Console.CursorSize = 50;
            Console.CursorTop = 20;
            Console.CursorVisible = false;
            Console.WriteLine("Main method complete. Press Enter.");
            Console.ReadLine();
        }
    }
}
