using System;

namespace Apress.VisualCSharpRecipes.Chapter01
{
    public class ConsoleUtils
    {
        // A method to display a prompt and read a response from the console.
        public static string ReadString(string msg)
        {
            Console.Write(msg);
            return Console.ReadLine();
        }

        // A method to display a message to the console.
        public static void WriteString(string msg)
        {
            Console.WriteLine(msg);
        }

        // Main method used for testing ConsoleUtility methods.
        public static void Main()
        {
            // Prompt the reader to enter a name.
            string name = ReadString("Please enter your name : ");

            // Welcome the reader to Visual C# 2010 Recipes.
            WriteString("Welcome to Visual C# 2010 Recipes, " + name);
        }
    }
}
