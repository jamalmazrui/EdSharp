using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Numerics;

namespace Apress.VisualCSharpRecipes.Chapter02
{
    class Recipe2_16
    {
        static void Main(string[] args)
        {
            // create a new big integer 
            BigInteger myBigInt = BigInteger.Multiply(Int64.MaxValue, 2);
            // add another value
            myBigInt = BigInteger.Add(myBigInt, Int64.MaxValue);
            // print out the value
            Console.WriteLine("Big Integer Value: {0}", myBigInt);

            // Wait to continue.
            Console.WriteLine("\n\nMain method complete. Press Enter");
            Console.ReadLine();
        }
    }
}
