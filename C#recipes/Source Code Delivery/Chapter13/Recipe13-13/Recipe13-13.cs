using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Apress.VisualCSharpRecipes.Chapter13
{
    class Recipe13_13
    {
        static void Main(string[] args)
        {
            // create an instance using eager initialization
            MyDataType eagerInstance = new MyDataType(false);
            Console.WriteLine("...do other things...");
            eagerInstance.sayHello();

            Lazy<MyDataType> lazyInstance = new Lazy<MyDataType>(() => new MyDataType(true));
            Console.WriteLine("...do other things...");
            lazyInstance.Value.sayHello();

            // Wait to continue.
            Console.WriteLine(Environment.NewLine);
            Console.WriteLine("Main method complete. Press Enter");
            Console.ReadLine();
        }
    }

    class MyDataType
    {
        public MyDataType(bool lazy)
        {
            Console.WriteLine("Initializing MyDataType - lazy instance: {0}", lazy);
        }

        public void sayHello()
        {
            Console.WriteLine("MyDataType Says Hello");
        }
    }
}
