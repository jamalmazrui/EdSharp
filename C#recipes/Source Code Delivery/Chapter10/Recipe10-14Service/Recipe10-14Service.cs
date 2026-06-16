using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.ServiceModel;

namespace Apress.VisualCSharpRecipes.Chapter10
{
    class Recipe10_14Service
    {
        static void Main(string[] args)
        {
            ServiceHost host = new ServiceHost(typeof(EmployeeService));
            host.Open();

            // Wait to continue.
            Console.WriteLine("Service host running. Press Enter to terminate.");
            Console.ReadLine();
        }
    }
}
