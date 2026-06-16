using System;

namespace Apress.VisualCSharpRecipes.Chapter01
{

    public class Recipe01_18
    {
        static void Main(string[] args)
        {

            // create an anoymous type           
            var joe = new {
                Name = "Joe Smith",
                Age  = 42,
                Family = new {
                    Father = "Pa Smith",
                    Mother = "Ma Smith",
                    Brother = "Pete Smith"
                },
            };

            // access the members of the anonymous type
            Console.WriteLine("Name: {0}", joe.Name);
            Console.WriteLine("Age: {0}", joe.Age);
            Console.WriteLine("Father: {0}", joe.Family.Father);
            Console.WriteLine("Mother: {0}", joe.Family.Mother);
            Console.WriteLine("Brother: {0}", joe.Family.Brother);

            Console.WriteLine("Main method complete. Press Enter.");
            Console.ReadLine();
        }
    }
}
