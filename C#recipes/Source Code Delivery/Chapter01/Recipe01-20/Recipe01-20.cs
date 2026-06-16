using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Dynamic;

namespace Apress.VisualCSharpRecipes.Chapter01
{

    public class Recipe01_20
    {
        // define a static string property
        static string MyStaticProperty
        {
            get;
            set;
        }


        static int MyStaticIntProperty
        {
            get;
            set;
        }

        static void Main(string[] args)
        {
            // write out the default values 
            Console.WriteLine("Default property values");
            Console.WriteLine("Default string property value: {0}", MyStaticProperty);
            Console.WriteLine("Default int property value: {0}", MyStaticIntProperty);
            
            // set the property values 
            MyStaticProperty = "Hello, World";
            MyStaticIntProperty = 32;

            // write out the changed values 
            Console.WriteLine("\nProperty values");
            Console.WriteLine("String property value: {0}", MyStaticProperty);
            Console.WriteLine("Int property value: {0}", MyStaticIntProperty);

            Console.WriteLine("\nMain method complete. Press Enter.");
            Console.ReadLine();
        }
    }
}
