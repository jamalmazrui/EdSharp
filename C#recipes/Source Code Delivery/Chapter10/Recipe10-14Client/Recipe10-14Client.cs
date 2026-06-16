using System;
using System.ServiceModel;

namespace Apress.VisualCSharpRecipes.Chapter10
{
    class Recipe10_14Client
    {
        static void Main(string[] args)
        {
            string serviceUri = "http://localhost:8000/EmployeeService";

            // Create the ChannelFactory that is used to generate service
            // proxies.
            using (ChannelFactory<IEmployeeService> channelFactory =
                new ChannelFactory<IEmployeeService>(new WSHttpBinding(), serviceUri))
            {
                // Create a dynamic proxy for IEmployeeService.
                IEmployeeService proxy = channelFactory.CreateChannel();

                // Create a new Employee.
                Employee emp = new Employee()
                {
                    DateOfBirth = DateTime.Now,
                    Name = "Allen Jones"
                };

                // Call the EmployeeService to create a new Employee record.
                emp = proxy.CreateEmployee(emp);

                Console.WriteLine("Created employee record - {0}", emp);

                // Update an existing Employee record
                emp.DateOfBirth = new DateTime(1950, 10, 13);
                emp = proxy.UpdateEmployee(emp);

                Console.WriteLine("Updated employee record - {0}", emp);
                
                // Wait to continue.
                Console.WriteLine(Environment.NewLine);
                Console.WriteLine("Main method complete. Press Enter");
                Console.ReadLine();
            }
        }
    }
}
