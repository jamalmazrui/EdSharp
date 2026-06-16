using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Apress.VisualCSharpRecipes.Chapter02
{
    class Recipe02_18
    {
        static void Main(string[] args)
        {

            // create a list of fruit, including duplicates
            List<Fruit> myList = new List<Fruit>() {
                new Fruit("apple", "green"),
                new Fruit("apple", "red"),
                new Fruit("orange", "orange"),
                new Fruit("orange", "orange"),
                new Fruit("banana", "yellow"),
                new Fruit("mango", "yellow"),
                new Fruit("cherry", "red"),
                new Fruit("fig", "brown"),
                new Fruit("fig", "brown"),
                new Fruit("fig", "brown"),
                new Fruit("cranberry", "red"),
                new Fruit("pear", "green")
            };

            // use the Distinct method to remove duplicates
            // and print out the unique entries that remain
            foreach (Fruit fruit in myList.Distinct(new FruitComparer()))
            {
                Console.WriteLine("Fruit: {0}:{1}", fruit.Name, fruit.Color);
            }

            // Wait to continue.
            Console.WriteLine("\n\nMain method complete. Press Enter");
            Console.ReadLine();
        }
    }

    class FruitComparer : IEqualityComparer<Fruit>
    {
        public bool Equals(Fruit first, Fruit second)
        {
            return first.Name == second.Name && first.Color == second.Color;
        }

        public int GetHashCode(Fruit fruit)
        {
            return fruit.Name.GetHashCode() + fruit.Name.GetHashCode();
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
