using System;
using System.Collections.Generic;

namespace Apress.VisualCSharpRecipes.Chapter13
{
    // The TeamMember class represents an individual team member.
    public class TeamMember
    {
        public string Name;
        public string Title;

        // Simple TeamMember constructor.
        public TeamMember(string name, string title)
        {
            Name = name;
            Title = title;
        }

        // Returns a string representation of the TeamMember.
        public override string ToString()
        {
            return string.Format("{0} ({1})", Name, Title);
        }
    }

    // Team class represents a collection of TeamMember objects.
    public class Team 
    {
        // A List to contain the TeamMember objects.
        private List<TeamMember> teamMembers = new List<TeamMember>();

        // Implement the GetEnumerator method, which will support
        // iteration across the entire team member List.
        public IEnumerator<TeamMember> GetEnumerator()
        {
            foreach (TeamMember tm in teamMembers)
            {
                yield return tm;
            }
        }

        // Implement the Reverse method, which will iterate through
        // the team members in alphabetical order.
        public IEnumerable<TeamMember> Reverse
        {
            get
            {
                for (int c = teamMembers.Count - 1; c >= 0; c--)
                {
                    yield return teamMembers[c];
                }
            }
        }

        // Implement the FirstTwo method, which will stop the iteration
        // after only the first two team members.
        public IEnumerable<TeamMember> FirstTwo
        {
            get
            {
                int count = 0;

                foreach (TeamMember tm in teamMembers)
                {
                    if (count >= 2)
                    {
                        // Terminate the iterator.
                        yield break;
                    }
                    else
                    {
                        // Return the TeamMember and maintain the iterator.
                        count++;
                        yield return tm;
                    }
                }
            }
        }

        // Adds a TeamMember object to the Team.
        public void AddMember(TeamMember member)
        {
            teamMembers.Add(member);
        }
    }

    // A class to demonstrate the use of Team.
    public class Recipe13_04
    {
        public static void Main()
        {
            // Create and populate a new Team.
            Team team = new Team();
            team.AddMember(new TeamMember("Curly", "Clown"));
            team.AddMember(new TeamMember("Nick", "Knife Thrower"));
            team.AddMember(new TeamMember("Nancy", "Strong Man"));

            // Enumerate the entire Team using the default iterator.
            Console.Clear();
            Console.WriteLine("Enumerate using default iterator:");
            foreach (TeamMember member in team)
            {
                Console.WriteLine("  " + member.ToString());
            }

            // Enumerate the first 2 Team members only.
            Console.WriteLine(Environment.NewLine); 
            Console.WriteLine("Enumerate using the FirstTwo iterator:");
            foreach (TeamMember member in team.FirstTwo)
            {
                Console.WriteLine("  " + member.ToString());
            }

            // Enumerate the entire Team in reverse order.
            Console.WriteLine(Environment.NewLine); 
            Console.WriteLine("Enumerate using the Reverse iterator:");
            foreach (TeamMember member in team.Reverse)
            {
                Console.WriteLine("  " + member.ToString());
            }

            // Wait to continue.
            Console.WriteLine(Environment.NewLine); 
            Console.WriteLine("Main method complete. Press Enter");
            Console.ReadLine();
        }
    }
}

