using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Apress.VisualCSharpRecipes.Chapter16
{
    class Recipe16_15
    {
        static void Main(string[] args)
        {
            // create the data sources
            IEnumerable<Fruit> datasource = createData();

            // perform a query 
            IEnumerable<Fruit> result = from e in datasource
                                        where e.ShelfLife <= 7
                                        select e;

            // enumerate the result elements
            Console.WriteLine("Results from enumeration");
            foreach (Fruit fruit in result)
            {
                Console.WriteLine("Name: {0} Color: {1} Shelf Life: {2} days.",
                    fruit.Name, fruit.Color, fruit.ShelfLife);
            }

            // convert the IEnumerable to an array
            Fruit[] array = result.ToArray<Fruit>();
            // print out the contents of the array
            Console.WriteLine("\nResults from array");
            for (int i = 0; i < array.Length; i++)
            {
                Console.WriteLine("Name: {0} Color: {1} Shelf Life: {2} days.",
                                    array[i].Name, array[i].Color, array[i].ShelfLife);
            }

            // convert the IEnumerable to a dictionary indexed by name
            Dictionary<string, Fruit> dictionary = result.ToDictionary(e => e.Name);
            // print out the contents of the dictionary
            Console.WriteLine("\nResults from dictionary");
            foreach (KeyValuePair<string, Fruit> kvp in dictionary)
            {
                Console.WriteLine("Name: {0} Color: {1} Shelf Life: {2} days.",
                    kvp.Key, kvp.Value.Color, kvp.Value.ShelfLife);
            }

            // convert the IEnumerable to a list
            IList<Fruit> list = result.ToList<Fruit>();
            // print out the contents of the list
            Console.WriteLine("\nResults from list");
            foreach (Fruit fruit in list)
            {
                Console.WriteLine("Name: {0} Color: {1} Shelf Life: {2} days.",
                    fruit.Name, fruit.Color, fruit.ShelfLife);
            }

            // convert the IEnumerable to a lookup, indexed by color
            ILookup<string, Fruit> lookup = result.ToLookup(e => e.Color);
            // print out the contents of the list
            Console.WriteLine("\nResults from lookup");
            IEnumerator<IGrouping<string, Fruit>> groups = lookup.GetEnumerator();
            while (groups.MoveNext())
            {
                IGrouping<string, Fruit> group = groups.Current;
                Console.WriteLine("Group for {0}", group.Key);
                foreach (Fruit fruit in group)
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
                new Fruit("apple", "red", 7),
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

