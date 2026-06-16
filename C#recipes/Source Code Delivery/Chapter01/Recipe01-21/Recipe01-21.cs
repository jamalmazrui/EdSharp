using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Dynamic;

namespace Apress.VisualCSharpRecipes.Chapter01
{
    public class Word
    {
        public string Text
        {
            get;
            set;
        }

        public static string operator +(Word w1, Word w2)
        {
            return w1.Text + " " + w2.Text;
        }

        public static Word operator +(Word w, int i)
        {
            return new Word() { Text = w.Text + i.ToString()};
        }

        public override string ToString()
        {
            return Text;
        }
    }

    public class Recipe01_21
    {
        static void Main(string[] args)
        {

            // create two word instances
            Word word1 = new Word() { Text = "Hello" };
            Word word2 = new Word() { Text = "World" };

            // print out the values
            Console.WriteLine("Word1: {0}", word1);
            Console.WriteLine("Word2: {0}", word2);
            Console.WriteLine("Added together: {0}", word1 + word2);
            Console.WriteLine("Added with int: {0}", word1 + 7);

            Console.WriteLine("\nMain method complete. Press Enter.");
            Console.ReadLine();
        }
    }
}
