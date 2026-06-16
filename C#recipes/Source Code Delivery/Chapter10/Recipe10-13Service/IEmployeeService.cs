using System;
using System.ServiceModel;

namespace Apress.VisualCSharpRecipes.Chapter10
{
    [ServiceContract(Namespace = "Apress.VisualCSharpRecipes.Chapter10")]
    public interface IEmployeeService
    {
        [OperationContract]
        Employee CreateEmployee(Employee newEmployee);

        [OperationContract]
        bool DeleteEmployee(int employeeId);

        [OperationContract(Name="GetEmployeeById")]
        Employee GetEmployee(int employeeId);

        [OperationContract(Name = "GetEmployeeByName")]
        Employee GetEmployee(string employeeName);

        [OperationContract]
        Employee UpdateEmployee(Employee updatedEmployee);
    }
}
