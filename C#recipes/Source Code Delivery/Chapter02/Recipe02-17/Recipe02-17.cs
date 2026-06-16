using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Apress.VisualCSharpRecipes.Chapter02
{
    class Recipe02_17
    {
        static void Main(string[] args)
        {
            // create a list of fruit
            List<Fruit> myList = new List<Fruit>() {
                new Fruit("apple", "green"),
                new Fruit("orange", "orange"),
                new Fruit("banana", "yellow"),
                new Fruit("mango", "yellow"),
                new Fruit("cherry", "red"),
                new Fruit("fig", "brown"),
                new Fruit("cranberry", "red"),
                new Fruit("pear", "green")
            };

            // select the names of fruit which isn't red and whose name
            // does not start with the letter 'c'
            IEnumerable<string> myResult = from e in myList where e.Color != "red" && e.Name[0] != 'c' orderby e.Name select e.Name;
            // write out the results
            foreach (string result in myResult)
            {
                Console.WriteLine("Result: {0}", result);
            }
            
            // perform the same query using lambda expressions
            myResult = myList.Where(e => e.Color != "red" && e.Name[0] != 'c').OrderBy(e => e.Name).Select(e => e.Name);
            // write out the results
            foreach (string result in myResult)
            {
                Console.WriteLine("Lambda Result: {0}", result);
            }

            // Wait to continue.
            Console.WriteLine("\n\nMain method complete. Press Enter");
            Console.ReadLine();
        }
    }

    class Fruit
    {
        public Fruit(string nameVal, string colorVal)
        {
            Name = nameVal;
            Color = colorVal;
        }
        public string Name { get; set; }
        public string Color { get; set; }
    }
}
