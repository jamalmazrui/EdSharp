using System;
using System.Reflection;
using System.Globalization;

namespace Apress.VisualCSharpRecipes.Chapter03
{
    class Recipe03_05
    {
        public static void ListAssemblies()
        {
            // Get an array of the assemblies loaded into the current 
            // application domain.
            Assembly[] assemblies = AppDomain.CurrentDomain.GetAssemblies();

            foreach (Assembly a in assemblies)
            {
                Console.WriteLine(a.GetName());
            }
        }
        
        public static void Main()
        {
            // List the assemblies in the current application domain.
            Console.WriteLine("**** BEFORE ****");
            ListAssemblies();

            // Load the System.Data assembly using a fully qualified display name.
            string name1 = "System.Data,Version=2.0.0.0," + 
                "Culture=neutral,PublicKeyToken=b77a5c561934e089";
            Assembly a1 = Assembly.Load(name1);

            // Load the System.Xml assembly using an AssemblyName.
            AssemblyName name2 = new AssemblyName();
            name2.Name = "System.Xml";
            name2.Version = new Version(2, 0, 0, 0);
            name2.CultureInfo = new CultureInfo("");    // Neutral culture.
            name2.SetPublicKeyToken(
                new byte[] {0xb7, 0x7a, 0x5c, 0x56, 0x19, 0x34, 0xe0, 0x89});
            Assembly a2 = Assembly.Load(name2);

            // Load the SomeAssembly assembly using a partial display name.
            Assembly a3 = Assembly.Load("SomeAssembly");

            // Load the assembly named c:\shared\MySharedAssembly.dll.
            Assembly a4 = Assembly.LoadFrom(@"c:\shared\MySharedAssembly.dll");

            // List the assemblies in the current application domain.
            Console.WriteLine("\n\n**** AFTER ****");
            ListAssemblies();

            // Wait to continue.
            Console.WriteLine("\nMain method complete. Press Enter.");
            Console.ReadLine();
        }
    }
}
