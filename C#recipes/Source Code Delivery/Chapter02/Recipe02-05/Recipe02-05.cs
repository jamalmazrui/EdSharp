using System;
using System.Text.RegularExpressions;

namespace Apress.VisualCSharpRecipes.Chapter02
{
    class Recipe02_05
    {
        /*
        // Alternative version of the ValidateInput method that does not create Regex instances.
        public static bool ValidateInput(string regex, string input)
        {
            // Test if the specified input matches the regular expression.
            return Regex.IsMatch(input, regex);
        }*/
        
        public static bool ValidateInput(string regex, string input)
        {
            // Create a new Regex based on the specified regular expression.
            Regex r = new Regex(regex);

            // Test if the specified input matches the regular expression.
            return r.IsMatch(input);
        }

        public static void Main(string[] args)
        {
            // Test the input from the command line. The first argument is the
            // regular expression, and the second is the input.
            Console.WriteLine("Regular Expression: {0}", args[0]);
            Console.WriteLine("Input: {0}", args[1]);
            Console.WriteLine("Valid = {0}", ValidateInput(args[0], args[1]));

            // Wait to continue.
            Console.WriteLine("\nMain method complete. Press Enter");
            Console.ReadLine();
        }
    }
}
