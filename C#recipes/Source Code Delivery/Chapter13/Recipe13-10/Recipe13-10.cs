using System;

namespace Apress.VisualCSharpRecipes.Chapter13
{
    public class SingletonExample
    {
        // A static member to hold a reference to the singleton instance.
        private static SingletonExample instance;

        // A static constructor to create the singleton instance. Another
        // alternative is to use lazy initialization in the Instance property.
        static SingletonExample()
        {
            instance = new SingletonExample();
        }

        // A private constructor to stop code from creating additional 
        // instances of the singleton type.
        private SingletonExample() { }

        // A public property to provide access to the singleton instance.
        public static SingletonExample Instance
        {
            get { return instance; }
        }

        // Public methods that provide singleton functionality.
        public void SomeMethod1() { /*..*/ }
        public void SomeMethod2() { /*..*/ }
    }

    // A class to demonstrate the use of SingletonExample.
    public class Recipe13_09
    {
        public static void Main()
        {
            // Obtain reference to singleton and invoke methods.
            SingletonExample s = SingletonExample.Instance;
            s.SomeMethod1();

            // Execute singleton functionality without a reference.
            SingletonExample.Instance.SomeMethod2();

            // Wait to continue.
            Console.WriteLine(Environment.NewLine);
            Console.WriteLine("Main method complete. Press Enter");
            Console.ReadLine();
        }
    }
}