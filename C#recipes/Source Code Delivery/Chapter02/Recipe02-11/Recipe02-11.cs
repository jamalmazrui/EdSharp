using System;
using System.Reflection;
using System.Collections.Generic;

namespace Apress.VisualCSharpRecipes.Chapter02
{
    class Recipe02_11
    {
        public static void Main(string[] args)
        {
            // Create an AssemblyName object for use during the example.
            AssemblyName assembly1 = new AssemblyName("com.microsoft.crypto, " +
                "Culture=en, PublicKeyToken=a5d015c7d5a0b012, Version=1.0.0.0");

            // Create and use a Dictionary of AssemblyName objects.
            Dictionary<string,AssemblyName> assemblyDictionary =
                new Dictionary<string,AssemblyName>();

            assemblyDictionary.Add("Crypto", assembly1);

            AssemblyName a1 = assemblyDictionary["Crypto"];

            Console.WriteLine("Got AssemblyName from dictionary: {0}", a1);

            // Create and use a List of Assembly Name objects.
            List<AssemblyName> assemblyList = new List<AssemblyName>();

            assemblyList.Add(assembly1);

            AssemblyName a2 = assemblyList[0];

            Console.WriteLine("\nFound AssemblyName in list: {0}", a1);

            // Create and use a Stack of Assembly Name objects
            Stack<AssemblyName> assemblyStack = new Stack<AssemblyName>();

            assemblyStack.Push(assembly1);

            AssemblyName a3 = assemblyStack.Pop();

            Console.WriteLine("\nPopped AssemblyName from stack: {0}", a1);

            // Wait to continue.
            Console.WriteLine("\nMain method complete. Press Enter");
            Console.ReadLine();
        }
    }
}
