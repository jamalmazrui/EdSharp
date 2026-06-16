using System;
using System.Text;
using System.Collections.Generic;

namespace Apress.VisualCSharpRecipes.Chapter13
{
    public class Employee : ICloneable
    {
        public string Name;
        public string Title;
        public int Age;

        // Simple Employee constructor.
        public Employee(string name, string title, int age)
        {
            Name = name;
            Title = title;
            Age = age;
        }

        // Create a clone using the Object.MemberwiseClone method because the
        // Employee class contains only string and value types.
        public object Clone()
        {
            return MemberwiseClone();
        }

        // Returns a string representation of the Employee object.
        public override string ToString()
        {
            return string.Format("{0} ({1}) - Age {2}", Name, Title, Age);
        }
    }

    public class Team : ICloneable
    {
        // A List to hold the Employee team members.
        public List<Employee> TeamMembers = 
            new List<Employee>();

        public Team()
        {
        }

        // Private constructor called by the Clone method to create a new Team
        // object and populate its List with clones of Employee objects from
        // a provided List.
        private Team(List<Employee> members)
        {
            foreach (Employee e in members)
            {
                // Clone the individual employee objects and
                // add them to the List.
                TeamMembers.Add((Employee)e.Clone());
            }
        }

        // Adds an Employee object to the Team.
        public void AddMember(Employee member)
        {
            TeamMembers.Add(member);
        }

        // Override Object.ToString to return a string representation of the
        // entire Team.
        public override string ToString()
        {
            StringBuilder str = new StringBuilder();

            foreach (Employee e in TeamMembers)
            {
                str.AppendFormat("  {0}\r\n", e);
            }

            return str.ToString();
        }

        // Implementation of ICloneable.Clone.
        public object Clone()
        {
            // Create a deep copy of the team by calling the private Team
            // constructor and passing the ArrayList containing team members.
            return new Team(this.TeamMembers);

            // The following command would create a shallow copy of the Team.
            // return MemberwiseClone();
        }
    }

    // A class to demonstrate the use of Employee.
    public class Recipe13_02
    {
        public static void Main()
        {
            // Create the original team.
            Team team = new Team();
            team.AddMember(new Employee("Frank", "Developer", 34));
            team.AddMember(new Employee("Kathy", "Tester", 78));
            team.AddMember(new Employee("Chris", "Support", 18));

            // Clone the original team.
            Team clone = (Team)team.Clone();

            // Display the original team.
            Console.WriteLine("Original Team:");
            Console.WriteLine(team);

            // Display the cloned team.
            Console.WriteLine("Clone Team:");
            Console.WriteLine(clone);

            // Make change.
            Console.WriteLine("*** Make a change to original team ***");
            Console.WriteLine(Environment.NewLine); 
            team.TeamMembers[0].Name = "Luke";
            team.TeamMembers[0].Title = "Manager";
            team.TeamMembers[0].Age = 44;

            // Display the original team.
            Console.WriteLine("Original Team:");
            Console.WriteLine(team);

            // Display the cloned team.
            Console.WriteLine("Clone Team:");
            Console.WriteLine(clone);

            // Wait to continue.
            Console.WriteLine(Environment.NewLine); 
            Console.WriteLine("Main method complete. Press Enter");
            Console.ReadLine();
        }
    }
}