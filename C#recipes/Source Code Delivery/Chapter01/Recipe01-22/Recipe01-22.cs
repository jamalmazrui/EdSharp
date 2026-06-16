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

        public static explicit operator Word(string str)
        {
            return new Word() { Text = str };
        }

        public static implicit operator string(Word w)
        {
            return w.Text;                
        }

        public static explicit operator int(Word w)
        {
            return w.Text.Length;
        }

        public override string ToString()
        {
            return Text;
        }
    }

    public class Recipe01_22
    {
        static void Main(string[] args)
        {

            // create a word instance
            Word word1 = new Word() { Text = "Hello"};

            // implicitly convert the word to a string
            string str1 = word1;
            // explicitly convert the word to a string
            string str2 = (string)word1;

            Console.WriteLine("{0} - {1}", str1, str2);

            // convert a string to a word
            Word word2 = (Word)"Hello";

            // convert a word to an int
            int count = (int)word2;

            Console.WriteLine("Length of {0} = {1}", word2.ToString(), count);	

            Console.WriteLine("\nMain method complete. Press Enter.");
            Console.ReadLine();
        }
    }
}
