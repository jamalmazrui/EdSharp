using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Reflection;

namespace Apress.VisualCSharpRecipes.Chapter03 
{
    class Recipe03_15
    {
        static void Main(string[] args)
        {
            // get the type we are interested in
            Type myType = typeof(Recipe03_15);

            // get the constructor details
            Console.WriteLine("\nConstructors...");
            foreach (ConstructorInfo constr in myType.GetConstructors())
            {
                Console.Write("Constructor: ");
                // get the paramters for this constructor
                foreach (ParameterInfo param in constr.GetParameters())
                {
                    Console.Write("{0} ({1}), ", param.Name, param.ParameterType);                        
                }
                Console.WriteLine();
            }

            // get the method details
            Console.WriteLine("\nMethods...");
            foreach (MethodInfo method in myType.GetMethods())
            {
                Console.Write(method.Name);
                // get the paramters for this constructor
                foreach (ParameterInfo param in method.GetParameters())
                {
                    Console.Write("{0} ({1}), ", param.Name, param.ParameterType);
                }
                Console.WriteLine();
            }

            // get the property details
            Console.WriteLine("\nProperty...");
            foreach (PropertyInfo property in myType.GetProperties())
            {
                Console.Write("{0} ", property.Name);
                // get the paramters for this constructor
                foreach (MethodInfo accessor in property.GetAccessors())
                {
                    Console.Write("{0}, ", accessor.Name);
                }
                Console.WriteLine();
            }
            
            // Wait to continue.
            Console.WriteLine("\nMain method complete. Press Enter.");
            Console.ReadLine();
        }

        public string MyProperty
        {
            get;
            set;
        }

        public Recipe03_15(string param1, int param2, char param3)
        {

        }
    }
}
