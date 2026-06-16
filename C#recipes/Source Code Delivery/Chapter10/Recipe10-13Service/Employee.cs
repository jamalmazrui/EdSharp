using System;
using System.Runtime.Serialization;

namespace Apress.VisualCSharpRecipes.Chapter10
{
    [DataContract]
    public class Employee
    {
        [DataMember]
        public DateTime DateOfBirth { get; set; }

        [DataMember]
        public int Id { get; set; }

        [DataMember]
        public string Name { get; set; }
    }
}
