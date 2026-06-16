using System;
using System.Collections.Generic;
using System.Linq;

namespace Apress.VisualCSharpRecipes.Chapter10
{
    public class EmployeeService : IEmployeeService
    {
        private Dictionary<int, Employee> employees;

        public EmployeeService()
        {
            employees = new Dictionary<int, Employee>();
        }

        // Create an Employee based on the contents of a provided
        // Employee object. Return the new Employee object.
        public Employee CreateEmployee(Employee newEmployee)
        {
            // NOTE: Should validate new employee data.
            newEmployee.Id = employees.Count + 1;

            lock (employees)
            {
                employees[newEmployee.Id] = newEmployee;
            }

            return newEmployee;
        }

        // Delete an employee by the specified Id and return true
        // or false depending on if en Employee was deleted.
        public bool DeleteEmployee(int employeeId)
        {
            lock(employees)
            {
                return employees.Remove(employeeId);
            }
        }

        // Get an employee by the specified Id and return null if
        // the employee does not exist.
        public Employee GetEmployee(int employeeId)
        {
            Employee employee = null;

            lock (employees)
            {
                employees.TryGetValue(employeeId, out employee);
            }

            return employee;
        }


        // Get an employee by the specified Name and return null if
        // the employee does not exist.
        public Employee GetEmployee(string employeeName)
        {
            Employee employee = null;

            lock (employees)
            {
                employee = employees.Values.FirstOrDefault
                    (e => e.Name.ToLower() == employeeName.ToLower());
            }

            return employee;
        }

        // Update an employee based on the contents of a provided
        // Employee object. Return the updated Employee object.
        public Employee UpdateEmployee(Employee updatedEmployee)
        {
            Employee employee = GetEmployee(updatedEmployee.Id);

            // NOTE: Should validate new employee data.
            if (employee != null)
            {
                lock (employees)
                {
                    employees[employee.Id] = updatedEmployee;
                }
            }

            return updatedEmployee;
        }
    }
}
