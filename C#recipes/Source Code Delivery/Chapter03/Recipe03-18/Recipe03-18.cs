using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Reflection;
using System.Dynamic;

namespace Apress.VisualCSharpRecipes.Chapter03
{
    class Recipe03_18
    {
        static void Main(string[] args)
        {
            dynamic dynamicDict = new MyDynamicDictionary();
            // set some properties
            Console.WriteLine("Setting property values");
            dynamicDict.FirstName = "Adam";
            dynamicDict.LastName = "Freeman";

            // get some properties
            Console.WriteLine("\nGetting property values");
            Console.WriteLine("Firstname {0}", dynamicDict.FirstName);
            Console.WriteLine("Lastname {0}", dynamicDict.LastName);

            // call an implemented member
            Console.WriteLine("\nGetting a static property");
            Console.WriteLine("Count {0}", dynamicDict.Count);

            Console.WriteLine("\nGetting a non-existent property");
            try
            {
                Console.WriteLine("City {0}", dynamicDict.City);
            }
            catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException e)
            {
                Console.WriteLine("Caught exception");
            }

            // Wait to continue.
            Console.WriteLine("\nMain method complete. Press Enter.");
            Console.ReadLine();
        }
    }
    class MyDynamicDictionary : DynamicObject
    {
        private IDictionary<string, object> dict = new Dictionary<string, object>();

        public int Count
        {
            get
            {
                Console.WriteLine("Get request for Count property");
                return dict.Count;
            }
        }

        public override bool TryGetMember(GetMemberBinder binder, out object result)
        {
            Console.WriteLine("Get request for {0}", binder.Name);
            return dict.TryGetValue(binder.Name, out result);
        }

        public override bool TrySetMember(SetMemberBinder binder, object value)
        {
            Console.WriteLine("Set request for {0}, value {1}", binder.Name, value);
            dict[binder.Name] = value;
            return true;
        }

    }
}
