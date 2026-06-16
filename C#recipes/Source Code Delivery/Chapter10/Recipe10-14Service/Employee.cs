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

        public override string ToString()
        {
            return String.Format("ID:{0}, NAME:{1}, DOB:{2}", Id, Name, DateOfBirth);
        }
    }
}
