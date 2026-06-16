using System;
using System.Collections.Generic;

namespace Apress.VisualCSharpRecipes.Chapter13
{
    public class Newspaper : IComparable<Newspaper>
    {
        private string name;
        private int circulation;

        private class AscendingCirculationComparer : IComparer<Newspaper>
        {
            // Implementation of IComparer.Compare. The generic definition of
            // IComparer allows us to ensure both arguments are Newspaper 
            // objects.
            public int Compare(Newspaper x, Newspaper y)
            {
                // Handle logic for null reference as dictated by the 
                // IComparer interface. Null is considered less than
                // any other value.
                if (x == null && y == null) return 0;
                else if (x == null) return -1;
                else if (y == null) return 1;

                // Short-circuit condition where x and y are references
                // to the same object.
                if (x == y) return 0;

                // Compare the circulation figures. IComparer dictates that:
                //      return less than zero if x < y
                //      return zero if x = y
                //      return greater than zero if x > y
                // This logic is easily implemented using integer arithmetic.
                return x.circulation - y.circulation;
            }
        }

        // Simple Newspaper constructor.
        public Newspaper(string name, int circulation)
        {
            this.name = name;
            this.circulation = circulation;
        }

        // Declare a read-only property that returns an instance of the 
        // AscendingCirculationComparer.
        public static IComparer<Newspaper> CirculationSorter
        {
            get { return new AscendingCirculationComparer(); }
        }

        // Override Object.ToString.
        public override string ToString()
        {
            return string.Format("{0}: Circulation = {1}", name, circulation);
        }

        // Implementation of IComparable.CompareTo. The generic definition
        // of IComparable allows us to ensure that the argument provided
        // must be a Newspaper object. Comparison is based on a  
        // case-insensitive comparison of the Newspaper names.
        public int CompareTo(Newspaper other)
        {
            // IComparable dictates that an object is always considered greater
            // than null.
            if (other == null) return 1;

            // Short-circuit the case where the other Newspaper object is a 
            // reference to this one.
            if (other == this) return 0;

            // Calculate return value by performing a case-insensitive 
            // comparison of the Newspaper names. 

            // Because the Newspaper name is a string, the easiest approach
            // is to rely on the comparison capabilities of the String 
            // class, which perform culture-sensitive string comparisons.
            return string.Compare(this.name, other.name, true);
        }
    }

    // A class to demonstrate the use of Newspaper.
    public class Recipe13_03
    {
        public static void Main()
        {
            List<Newspaper> newspapers = new List<Newspaper>();

            newspapers.Add(new Newspaper("The Echo", 125780));
            newspapers.Add(new Newspaper("The Times", 55230));
            newspapers.Add(new Newspaper("The Gazette", 235950));
            newspapers.Add(new Newspaper("The Sun", 88760));
            newspapers.Add(new Newspaper("The Herald", 5670));

            Console.Clear();
            Console.WriteLine("Unsorted newspaper list:");
            foreach (Newspaper n in newspapers)
            {
                Console.WriteLine("  " + n);
            }

            Console.WriteLine(Environment.NewLine); 
            Console.WriteLine("Newspaper list sorted by name (default order):");
            newspapers.Sort();
            foreach (Newspaper n in newspapers)
            {
                Console.WriteLine("  " + n);
            }

            Console.WriteLine(Environment.NewLine); 
            Console.WriteLine("Newspaper list sorted by circulation:");
            newspapers.Sort(Newspaper.CirculationSorter);
            foreach (Newspaper n in newspapers)
            {
                Console.WriteLine("  " + n);
            }

            // Wait to continue.
            Console.WriteLine(Environment.NewLine); 
            Console.WriteLine("Main method complete. Press Enter");
            Console.ReadLine();
        }
    }
}