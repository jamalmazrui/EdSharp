using System;
using System.Text;

namespace Apress.VisualCSharpRecipes.Chapter03 {
    class Recipe03_10 {
        public static void Main() {
            // Obtain type information using the typeof operator.
            Type t1 = typeof(StringBuilder);
            Console.WriteLine(t1.AssemblyQualifiedName);

            // Obtain type information using the Type.GetType method.
            // Case sensitive, return null if not found.
            Type t2 = Type.GetType("System.String");
            Console.WriteLine(t2.AssemblyQualifiedName);

            // Case-sensitive, throw TypeLoadException if not found.
            Type t3 = Type.GetType("System.String", true);
            Console.WriteLine(t3.AssemblyQualifiedName);

            // Case-insensitive, throw TypeLoadException if not found.
            Type t4 = Type.GetType("system.string", true, true);
            Console.WriteLine(t4.AssemblyQualifiedName);

            // Assembly-qualifed type name.
            Type t5 = Type.GetType("System.Data.DataSet,System.Data," +
                "Version=2.0.0.0,Culture=neutral,PublicKeyToken=b77a5c561934e089");
            Console.WriteLine(t5.AssemblyQualifiedName);

            // Obtain type information using the Object.GetType method.
            StringBuilder sb = new StringBuilder();
            Type t6 = sb.GetType();
            Console.WriteLine(t6.AssemblyQualifiedName);

            // Wait to continue.
            Console.WriteLine("\nMain method complete. Press Enter.");
            Console.ReadLine();
        }
    }
}
