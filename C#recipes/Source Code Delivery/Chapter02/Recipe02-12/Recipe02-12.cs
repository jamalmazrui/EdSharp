using System;
using System.Collections.Generic;

namespace Apress.VisualCSharpRecipes.Chapter02
{
    public class Bag<T>
    {
        // A List to hold the bags's contents. The list must be 
        // of the same type as the bag.
        private List<T> items = new List<T>();

        // A method to add an item to the bag.
        public void Add(T item)
        {
            items.Add(item);
        }

        // A method to get a random item from the bag.
        public T Remove()
        {
            T item = default(T);

            if (items.Count != 0)
            {
                // Determine which item to remove from the bag.
                Random r = new Random();
                int num = r.Next(0, items.Count);

                // Remove the item
                item = items[num];
                items.RemoveAt(num);
            }
            return item;
        }

        // a method to provide an enumerator from the underlying list
        public IEnumerator<T> GetEnumerator()
        {
            return items.GetEnumerator();
        }

        // A method to remove all items from the bag and return them
        // as an array.
        public T[] RemoveAll()
        {
            T[] i = items.ToArray();
            items.Clear();
            return i;
        }
    }

    public class Recipe02_12
    {
        public static void Main(string[] args)
        {
            // Create a new bag of strings.
            Bag<string> bag = new Bag<string>();

            // Add strings to the bag.
            bag.Add("Darryl");
            bag.Add("Bodders");
            bag.Add("Gary");
            bag.Add("Mike");
            bag.Add("Nigel");
            bag.Add("Ian");

            Console.WriteLine("Bag contents are:");
            foreach (string elem in bag)
            {
                Console.WriteLine("Element: {0}", elem);
            }

            // Take four strings from the bag and display.
            Console.WriteLine("\nRemoving individual elements");
            Console.WriteLine("Removing = {0}", bag.Remove());
            Console.WriteLine("Removing = {0}", bag.Remove());
            Console.WriteLine("Removing = {0}", bag.Remove());
            Console.WriteLine("Removing = {0}", bag.Remove());

            Console.WriteLine("\nBag contents are:");
            foreach (string elem in bag)
            {
                Console.WriteLine("Element: {0}", elem);
            }

            // Remove the remaining items from the bag. 
            Console.WriteLine("\nRemoving all elements");
            string[] s = bag.RemoveAll();

            Console.WriteLine("\nBag contents are:");
            foreach (string elem in bag)
            {
                Console.WriteLine("Element: {0}", elem);
            }

            // Wait to continue.
            Console.WriteLine("\nMain method complete. Press Enter");
            Console.ReadLine();
        }
    }
}
