using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Apress.VisualCSharpRecipes.Chapter16
{
    class Recipe16_05
    {
        static void Main(string[] args)
        {
            IList<Fruit> sourcedata = createData();
            var result = from e in sourcedata
                         select new
                         {
                             e.Name,
                             e.Color
                         };
            foreach (var element in result)
            {
                Console.WriteLine("Result: {0} {1}", element.Name, element.Color);
            }

            // Wait to continue.
            Console.WriteLine("\nMain method complete. Press Enter");
            Console.ReadLine(); 
        }

        static IList<Fruit> createData()
        {
            return new List<Fruit>()
            {
                new Fruit("apple", "green", 7),
                new Fruit("orange", "orange", 10),
                new Fruit("grape", "green", 4),
                new Fruit("fig", "brown", 12),
                new Fruit("plum", "red", 2),
                new Fruit("banana", "yellow", 10),
                new Fruit("cherry", "red", 7)
            };
        }
    }
    class Fruit
    {
        public Fruit(string namearg, string colorarg, int lifearg)
        {
            Name = namearg;
            Color = colorarg;
            ShelfLife = lifearg;
        }
        public string Name { get; set; }
        public string Color { get; set; }
        public int ShelfLife { get; set; }
    }
}
