using System;
using System.IO;

namespace Apress.VisualCSharpRecipes.Chapter03
{
    class Recipe03_11
    {
        // A method to test if an object is an instance of a type or a derived type.
        public static bool IsType(object obj, string type) 
        {
            // Get the named type, use case insensitive search, throw 
            // an exception if the type is not found.
            Type t = Type.GetType(type, true, true);

            return t == obj.GetType() || obj.GetType().IsSubclassOf(t);
        }

        public static void Main() 
        {
            // Create a new StringReader for testing.
            Object someObject = new StringReader("This is a StringReader");

            // Test if someObject is a StringReader by obtaining and 
            // comparing a Type reference using the typeof operator.
            if (typeof(StringReader) == someObject.GetType()) 
            {
                Console.WriteLine("typeof: someObject is a StringReader");
            }

            // Test if someObject is, or is derived from, a TextReader
            // using the is operator.
            if (someObject is TextReader) 
            {
                Console.WriteLine(
                    "is: someObject is a TextReader or a derived class");
            }

            // Test if someObject is, or is derived from, a TextReader using 
            // the Type.GetType and Type.IsSubclassOf methods.
            if (IsType(someObject, "System.IO.TextReader")) 
            {
                Console.WriteLine("GetType: someObject is a TextReader");
            }

            // Use the "as" operator to perform a safe cast.
            StringReader reader = someObject as StringReader;
            if (reader != null) 
            {
                Console.WriteLine("as: someObject is a StringReader");
            }

            // Wait to continue.
            Console.WriteLine("\nMain method complete. Press Enter.");
            Console.ReadLine();
        }
    }
}