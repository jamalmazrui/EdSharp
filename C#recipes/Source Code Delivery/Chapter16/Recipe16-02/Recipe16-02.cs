using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Xml.Linq;

namespace Apress.VisualCSharpRecipes.Chapter16
{
    class Recipe16_02
    {
        static void Main(string[] args)
        {
            // create the data
            IList<Fruit> datasource = createData();

            // filter based on a single characteristic
            IEnumerable<string> result1 = from e in datasource 
                                          where e.Color == "green" select e.Name;
            Console.WriteLine("Filter for green fruit");
            foreach (string str in result1)
            {
                Console.WriteLine("Fruit {0}", str);
            }

            // filter based using > operator
            IEnumerable<Fruit> result2 = from e in datasource 
                                         where e.ShelfLife > 5 select e;
            Console.WriteLine("\nFilter for life > 5 days");
            foreach (Fruit fruit in result2)
            {
                Console.WriteLine("Fruit {0}", fruit.Name);
            }

            // filter using two characteristics
            IEnumerable<string> result3 = from e in datasource 
                                          where e.Color == "green" && e.ShelfLife > 5
                                          select e.Name;
            Console.WriteLine("\nFilter for green fruit and life > 5 days");
            foreach (string str in result3)
            {
                Console.WriteLine("Fruit {0}", str);
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
        public string Name { get; set;}
        public string Color { get; set;}
        public int ShelfLife { get; set;}
    }

}
