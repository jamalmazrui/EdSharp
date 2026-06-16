using System;
using System.ServiceModel;

namespace Apress.VisualCSharpRecipes.Chapter10
{
    public static class Recipe10_13Service
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
