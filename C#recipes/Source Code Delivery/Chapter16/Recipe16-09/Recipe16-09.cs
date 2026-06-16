using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Apress.VisualCSharpRecipes.Chapter16
{
    class Recipe16_09
    {
        static void Main(string[] args)
        {
            // create the data source
            IList<Fruit> datasource = createData();

            Console.WriteLine("Perfoming group...by... query");
            // perform a query with a basic grouping
            IEnumerable<IGrouping<string, Fruit>> result =
                from e in datasource group e by e.Color;

            foreach (IGrouping<string, Fruit> group in result)
            {
                Console.WriteLine("\nStart of group for {0}", group.Key);
                foreach (Fruit fruit in group)
                {
                    Console.WriteLine("Name: {0} Color: {1} Shelf Life: {2} days.",
                        fruit.Name, fruit.Color, fruit.ShelfLife);
                }
            }

            Console.WriteLine("\n\nPerfoming group...by...into query");
            // use the group into keywords
            var result2 = from e in datasource
                         group e by e.Color into g
                         select new
                         {
                            Color = g.Key,
                            Fruits = g
                         };

            foreach (var element in result2)
            {
                Console.WriteLine("\nElement for color {0}", element.Color);
                foreach (var fruit in element.Fruits)
                {
                    Console.WriteLine("Name: {0} Color: {1} Shelf Life: {2} days.",
                        fruit.Name, fruit.Color, fruit.ShelfLife);
                }
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
