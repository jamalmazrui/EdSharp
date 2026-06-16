using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Apress.VisualCSharpRecipes.Chapter02
{
    class Recipe02_15
    {
        static void Main(string[] args)
        {
            // Local varable to hold the key entered by the user.
            ConsoleKeyInfo key;

            // Control whether character or asterisk is displayed.
            bool secret = false;

            // Character List for the user data entered.
            List<char> input = new List<char>();

            string msg = "Enter characters and press Escape to see input." +
                "\nPress F1 to enter/exit Secret mode and Alt-X to exit.";

            Console.WriteLine(msg);

            // Process input until the user enters "Alt-X" or "Alt-x".
            do
            {
                // Read a key from the Console. Intercept the key so that it is not 
                // displayed to the Console. What is displayed is determined later 
                // depending on whether the program is in secret mode or not.
                key = Console.ReadKey(true);

                // Switch secret mode on and off.
                if (key.Key == ConsoleKey.F1)
                {
                    if (secret)
                    {
                        // Switch secret mode off.
                        secret = false;
                    }
                    else
                    {
                        // Switch secret mode on.
                        secret = true;
                    }
                }

                // Handle backspace.
                if (key.Key == ConsoleKey.Backspace)
                {
                    if (input.Count > 0)
                    {
                        // Backspace pressed, remove the last character.
                        input.RemoveAt(input.Count - 1);

                        Console.Write(key.KeyChar);
                        Console.Write(" ");
                        Console.Write(key.KeyChar);
                    }
                }
                // Handle Escape.
                else if (key.Key == ConsoleKey.Escape)
                {
                    Console.Clear();
                    Console.WriteLine("Input: {0}\n\n",
                        new String(input.ToArray()));
                    Console.WriteLine(msg);
                    input.Clear();
                }
                // Handle character input.
                else if (key.Key >= ConsoleKey.A && key.Key <= ConsoleKey.Z)
                {
                    input.Add(key.KeyChar);
                    if (secret)
                    {
                        Console.Write("*");
                    }
                    else
                    {
                        Console.Write(key.KeyChar);
                    }
                }
            } while (key.Key != ConsoleKey.X
                || key.Modifiers != ConsoleModifiers.Alt);

            // Wait to continue.
            Console.WriteLine("\n\nMain method complete. Press Enter");
            Console.ReadLine();
        }        
    }
}
