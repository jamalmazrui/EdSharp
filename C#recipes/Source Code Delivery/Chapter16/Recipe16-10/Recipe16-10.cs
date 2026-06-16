using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Apress.VisualCSharpRecipes.Chapter16
{
    class Recipe16_10
    {
        static void Main(string[] args)
        {
            // create the data source
            IList<Fruit> datasource = createData();

            IEnumerable<Fruit> result = from e in datasource
                                        orderby e.Name
                                        orderby e.Color descending
                                        select e;

            foreach (Fruit fruit in result)
            {
                Console.WriteLine("Name: {0} Color: {1} Shelf Life: {2} days.",
                    fruit.Name, fruit.Color, fruit.ShelfLife);
            }

            // Wait to continue.
            Console.WriteLine("\nMain method complete. Press Enter");
            Console.ReadLine();
        }

        static IList<Fruit> createData()
        {
            return new List<Fruit>()
            {
                new Fruit("apple", "red", 7),
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
