using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Dynamic;

namespace Apress.VisualCSharpRecipes.Chapter01
{
    public class Recipe01_19
    {
        static void Main(string[] args)
        {
            // create the expando
            dynamic expando = new ExpandoObject();
            expando.Name = "Joe Smith";
            expando.Age = 42;
            expando.Family = new ExpandoObject();
            expando.Family.Father = "Pa Smith";
            expando.Family.Mother = "Ma Smith";
            expando.Family.Brother = "Pete Smith";

            // access the members of the dynamic type
            Console.WriteLine("Name: {0}", expando.Name);
            Console.WriteLine("Age: {0}", expando.Age);
            Console.WriteLine("Father: {0}", expando.Family.Father);
            Console.WriteLine("Mother: {0}", expando.Family.Mother);
            Console.WriteLine("Brother: {0}", expando.Family.Brother);

            // change a value
            expando.Age = 44;
            // add a new property
            expando.Family.Sister = "Kathy Smith";

            Console.WriteLine("\nModified Values");
            Console.WriteLine("Age: {0}", expando.Age);
            Console.WriteLine("Sister: {0}", expando.Family.Sister);

            Console.WriteLine("Main method complete. Press Enter.");
            Console.ReadLine();
        }
    }
}
