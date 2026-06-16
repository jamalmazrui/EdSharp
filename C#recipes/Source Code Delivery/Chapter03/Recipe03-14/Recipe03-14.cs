using System;

namespace Apress.VisualCSharpRecipes.Chapter03
{
    [Author("Lena")]
    [Author("Allen", Company = "Apress")]
    class Recipe03_14
    {
        public static void Main()
        {
            // Get a Type object for this class. 
            Type type = typeof(Recipe03_14);

            // Get the attributes for the type. Apply a filter so that only
            // instances of AuthorAttribute are returned.
            object[] attrs =
                type.GetCustomAttributes(typeof(AuthorAttribute), true);

            // Enumerate the attributes and display their details.
            foreach (AuthorAttribute a in attrs)
            {
                Console.WriteLine(a.Name + ", " + a.Company);
            }

            // Wait to continue.
            Console.WriteLine("\nMain method complete. Press Enter.");
            Console.ReadLine();
        }
    }
}
