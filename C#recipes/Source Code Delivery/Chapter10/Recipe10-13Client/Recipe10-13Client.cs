using System;
using Apress.VisualCSharpRecipes.Chapter10.Services;

namespace Apress.VisualCSharpRecipes.Chapter10
{
    class Recipe10_13Client
    {
        private static string FormatEmployee(Employee emp)
        {
            return String.Format("ID:{0}, NAME:{1}, DOB:{2}", 
                emp.Id, emp.Name, emp.DateOfBirth);
        }

        static void Main(string[] args)
        {
            // Create a service proxy.
            using (EmployeeServiceClient employeeService
                = new EmployeeServiceClient())
            {
                // Create a new Employee.
                Employee emp = new Employee()
                {
                    DateOfBirth = DateTime.Now,
                    Name = "Allen Jones"
                };

                // Call the EmployeeService to create a new Employee record.
                emp = employeeService.CreateEmployee(emp);

                Console.WriteLine("Created employee record - {0}", 
                    FormatEmployee(emp));

                // Update the existing Employee.
                emp.DateOfBirth = new DateTime(1950, 10, 13);
                emp = employeeService.UpdateEmployee(emp);

                Console.WriteLine("Updated employee record - {0}", 
                    FormatEmployee(emp));

                // Wait to continue.
                Console.WriteLine(Environment.NewLine);
                Console.WriteLine("Main method complete. Press Enter");
                Console.ReadLine();
            }
        }
    }
}
