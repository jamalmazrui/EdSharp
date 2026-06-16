using System;

namespace Apress.VisualCSharpRecipes.Chapter01
{
    public static class MyStaticClass
    {
        public static string getMessage()
        {
            return "This is a static member";
        }

        public static string StaticProperty
        {
            get;
            set;
        }

    }

    public class Recipe01_16
    {
        static void Main(string[] args)
        {
            // call the static method and print out the result
            Console.WriteLine(MyStaticClass.getMessage());

            // set and get the property
            MyStaticClass.StaticProperty = "this is the property value";
            Console.WriteLine(MyStaticClass.StaticProperty);
            
            Console.WriteLine("Main method complete. Press Enter.");
            Console.ReadLine();
        }
    }
}
