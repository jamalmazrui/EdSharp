using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Apress.VisualCSharpRecipes.Chapter16
{
    class Recipe16_06
    {
        static void Main(string[] args)
        {
            // create the data sources
            IList<FruitColor> colorsource = createColorData();
            IList<FruitShelfLife> shelflifesource = createShelfLifeData();

            // perform the LINQ query with a join
            var result = from e in colorsource
                         join f in shelflifesource on e.Name equals f.Name
                         where e.Color == "green"
                         select new
                         {
                             e.Name,
                             e.Color,
                             f.Life
                         };

            // write out the results
            foreach (var element in result)
            {
                Console.WriteLine("Name: {0} Color {1} Shelf Life: {2} days",
                    element.Name, element.Color, element.Life);
            }

            // Wait to continue.
            Console.WriteLine("\nMain method complete. Press Enter");
            Console.ReadLine(); 
        }

        static IList<FruitColor> createColorData()
        {
            return new List<FruitColor>()
            {
                new FruitColor("apple", "green"),
                new FruitColor("orange", "orange"),
                new FruitColor("grape", "green"),
                new FruitColor("fig", "brown"),
                new FruitColor("plum", "red"),
                new FruitColor("banana", "yellow"),
                new FruitColor("cherry", "red")
            };
        }

        static IList<FruitShelfLife> createShelfLifeData()
        {
            return new List<FruitShelfLife>()
            {
                new FruitShelfLife("apple", 7),
                new FruitShelfLife("orange", 10),
                new FruitShelfLife("grape", 4),
                new FruitShelfLife("fig", 12),
                new FruitShelfLife("plum", 2),
                new FruitShelfLife("banana", 10),
                new FruitShelfLife("cherry",  7)
            };
        }
    }

    class FruitColor
    {
        public FruitColor(string namearg, string colorarg)
        {
            Name = namearg;
            Color = colorarg;
        }
        public string Name { get; set; }
        public string Color { get; set; }
    }

    class FruitShelfLife
    {
        public FruitShelfLife(string namearg, int lifearg)
        {
            Name = namearg;
            Life = lifearg;
        }
        public string Name { get; set; }
        public int Life{ get; set; }
    }    
}
