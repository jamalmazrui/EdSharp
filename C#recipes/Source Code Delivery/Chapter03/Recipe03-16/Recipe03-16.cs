using System;
using System.Reflection;

namespace Apress.VisualCSharpRecipes.Chapter03
{
    class Recipe03_16
    {
        static void Main(string[] args)
        {
            // create an instance of this type
            object myInstance = new Recipe03_16();

            // get the type we are interested in
            Type myType = typeof(Recipe03_16);

            // get the method information
            MethodInfo methodInfo = myType.GetMethod("printMessage", 
                new Type[] { typeof(string), typeof(int), typeof(char) });
            // invoke the method through the methodinfo
            methodInfo.Invoke(myInstance, new object[] { "hello", 37, 'c' });

            // invoke the method using the instance we created
            myType.InvokeMember("printMessage", BindingFlags.InvokeMethod, null, myInstance,
                new object[] { "hello", 37, 'c' });

            methodInfo.Invoke(null, BindingFlags.InvokeMethod, null,
                new object[] { "hello", 37, 'c' }, null);

            // Wait to continue.
            Console.WriteLine("\nMain method complete. Press Enter.");
            Console.ReadLine();
        }

        public static void printMessage(string param1, int param2, char param3)
        {
            Console.WriteLine("PrintMessage {0} {1} {2}", param1, param2, param3);
        }
    }
}
