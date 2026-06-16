using System;
using System.IO;
using System.Text;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Formatters.Binary;

namespace Apress.VisualCSharpRecipes.Chapter13
{
    [Serializable]
    public class Employee : ISerializable
    {
        private string name;
        private int age;
        private string address;

        // Simple Employee constructor.
        public Employee(string name, int age, string address)
        {
            this.name = name;
            this.age = age;
            this.address = address;
        }

        // Constructor required to enable a formatter to deserialize an 
        // Employee object. You should declare the constructor private or at
        // least protected to ensure it is not called unnecessarily.
        private Employee(SerializationInfo info, StreamingContext context)
        {
            // Extract the name and age of the Employee, which will always be 
            // present in the serialized data regardless of the value of the
            // StreamingContext.
            name = info.GetString("Name");
            age = info.GetInt32("Age");

            // Attempt to extract the Employee's address and fail gracefully
            // if it is not available. 
            try
            {
                address = info.GetString("Address");
            }
            catch (SerializationException)
            {
                address = null;
            }
        }

        // Public property to provide access to employee's name.
        public string Name
        {
            get { return name; }
            set { name = value; }
        }

        // Public property to provide access to employee's age.
        public int Age
        {
            get { return age; }
            set { age = value; }
        }

        // Public property to provide access to employee's address.
        // Uses lazy initialization to establish address because
        // a deserialized object will not have an address value.
        public string Address
        {
            get
            {
                if (address == null)
                {
                    // Load the address from persistent storage.
                    // In this case, set it to an empty string.
                    address = String.Empty;
                }
                return address;
            }

            set
            {
                address = value;
            }
        }

        // Declared by the ISerializable interface, the GetObjectData method
        // provides the mechanism with which a formatter obtains the object
        // data that it should serialize.
        public void GetObjectData(SerializationInfo inf, StreamingContext con)
        {
            // Always serialize the Employee's name and age.
            inf.AddValue("Name", name);
            inf.AddValue("Age", age);

            // Don't serialize the Employee's address if the StreamingContext 
            // indicates that the serialized data is to be written to a file.
            if ((con.State & StreamingContextStates.File) == 0)
            {
                inf.AddValue("Address", address);
            }
        }

        // Override Object.ToString to return a string representation of the
        // Employee state.
        public override string ToString()
        {
            StringBuilder str = new StringBuilder();

            str.AppendFormat("Name: {0}\r\n", Name);
            str.AppendFormat("Age: {0}\r\n", Age);
            str.AppendFormat("Address: {0}\r\n", Address);

            return str.ToString();
        }
    }

    // A class to demonstrate the use of Employee.
    public class Recipe13_01
    {
        public static void Main(string[] args)
        {
            // Create an Employee object representing Roger.
            Employee roger = new Employee("Roger", 56, "London");

            // Display Roger.
            Console.WriteLine(roger);

            // Serialize Roger specifying another application domain as the 
            // destination of the serialized data. All data including Roger's
            // address is serialized.
            Stream str = File.Create("roger.bin");
            BinaryFormatter bf = new BinaryFormatter();
            bf.Context = 
                new StreamingContext(StreamingContextStates.CrossAppDomain);
            bf.Serialize(str, roger);
            str.Close();

            // Deserialize and display Roger.
            str = File.OpenRead("roger.bin");
            bf = new BinaryFormatter();
            roger = (Employee)bf.Deserialize(str);
            str.Close();
            Console.WriteLine(roger);

            // Serialize Roger specifying a file as the destination of the 
            // serialized data. In this case, Roger's address is not included
            // in the serialized data.
            str = File.Create("roger.bin");
            bf = new BinaryFormatter();
            bf.Context = new StreamingContext(StreamingContextStates.File);
            bf.Serialize(str, roger);
            str.Close();

            // Deserialize and display Roger.
            str = File.OpenRead("roger.bin");
            bf = new BinaryFormatter();
            roger = (Employee)bf.Deserialize(str);
            str.Close();
            Console.WriteLine(roger);

            // Wait to continue.
            Console.WriteLine(Environment.NewLine);
            Console.WriteLine("Main method complete. Press Enter");
            Console.ReadLine();
        }
    }
}