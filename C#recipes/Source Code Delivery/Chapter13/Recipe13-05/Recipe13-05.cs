using System;
using System.Collections;

namespace Apress.VisualCSharpRecipes.Chapter13
{
    // TeamMember class represents an individual team member.
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

    // Team class represents a collection of TeamMember objects. Implements
    // the IEnumerable interface to support enumerating TeamMember objects.
    public class Team : IEnumerable
    {
        // TeamMemberEnumerator is a private nested class that provides
        // the functionality to enumerate the TeamMembers contained in 
        // a Team collection. As a nested class, TeamMemberEnumerator
        // has access to the private members of the Team class.
        private class TeamMemberEnumerator : IEnumerator
        {
            // The Team that this object is enumerating.
            private Team sourceTeam;

            // Boolean to indicate whether underlying Team has changed
            // and so is invalid for further enumeration.
            private bool teamInvalid = false;

            // Integer to identify the current TeamMember. Provides
            // the index of the TeamMember in the underlying ArrayList
            // used by the Team collection. Initialize to -1, which is 
            // the index prior to the first element.
            private int currentMember = -1;

            // Constructor takes a reference to the Team that is the source
            // of enumerated data.
            internal TeamMemberEnumerator(Team team)
            {
                this.sourceTeam = team;

                // Register with sourceTeam for change notifications.
                sourceTeam.TeamChange +=
                    new TeamChangedEventHandler(this.TeamChange);
            }

            // Implement the IEnumerator.Current property.
            public object Current
            {
                get
                {
                    // If the TeamMemberEnumerator is positioned before 
                    // the first element or after the last element, then 
                    // throw an exception.
                    if (currentMember == -1 ||
                        currentMember > (sourceTeam.teamMembers.Count - 1))
                    {
                        throw new InvalidOperationException();
                    }

                    //Otherwise, return the current TeamMember.
                    return sourceTeam.teamMembers[currentMember];
                }
            }

            // Implement the IEnumerator.MoveNext method.
            public bool MoveNext()
            {
                // If underlying Team is invalid, throw exception.
                if (teamInvalid)
                {
                    throw new InvalidOperationException("Team modified");
                }

                // Otherwise, progress to the next TeamMember.
                currentMember++;

                // Return false if we have moved past the last TeamMember.
                if (currentMember > (sourceTeam.teamMembers.Count - 1))
                {
                    return false;
                }
                else
                {
                    return true;
                }
            }

            // Implement the IEnumerator.Reset method.
            // This method resets the position of the TeamMemberEnumerator
            // to the beginning of the TeamMembers collection.
            public void Reset()
            {
                // If underlying Team is invalid, throw exception.
                if (teamInvalid)
                {
                    throw new InvalidOperationException("Team modified");
                }

                // Move the currentMember pointer back to the index
                // preceding the first element.
                currentMember = -1;
            }

            // An event handler to handle notifications that the underlying
            // Team collection has changed.
            internal void TeamChange(Team t, EventArgs e)
            {
                // Signal that the underlying Team is now invalid.
                teamInvalid = true;
            }
        }

        // A delegate that specifies the signature that all team change event
        // handler methods must implement.
        public delegate void TeamChangedEventHandler(Team t, EventArgs e);

        // An ArrayList to contain the TeamMember objects.
        private ArrayList teamMembers;

        // The event used to notify TeamMemberEnumerators that the Team
        // has changed. 
        public event TeamChangedEventHandler TeamChange;

        // Team constructor.
        public Team()
        {
            teamMembers = new ArrayList();
        }

        // Implement the IEnumerable.GetEnumerator method.
        public IEnumerator GetEnumerator()
        {
            return new TeamMemberEnumerator(this);
        }

        // Adds a TeamMember object to the Team.
        public void AddMember(TeamMember member)
        {
            teamMembers.Add(member);

            // Notify listeners that the list has changed.
            if (TeamChange != null)
            {
                TeamChange(this, null);
            }
        }
    }

    // A class to demonstrate the use of Team.
    public class Recipe13_05
    {
        public static void Main()
        {
            // Create a new Team.
            Team team = new Team();
            team.AddMember(new TeamMember("Curly", "Clown"));
            team.AddMember(new TeamMember("Nick", "Knife Thrower"));
            team.AddMember(new TeamMember("Nancy", "Strong Man"));

            // Enumerate the Team.
            Console.Clear();
            Console.WriteLine("Enumerate with foreach loop:");
            foreach (TeamMember member in team)
            {
                Console.WriteLine(member.ToString());
            }

            // Enumerate using a While loop.
            Console.WriteLine(Environment.NewLine); 
            Console.WriteLine("Enumerate with while loop:");
            IEnumerator e = team.GetEnumerator();
            while (e.MoveNext())
            {
                Console.WriteLine(e.Current);
            }

            // Enumerate the Team and try to add a Team Member.
            // (This will cause an exception to be thrown.)
            Console.WriteLine(Environment.NewLine);
            Console.WriteLine("Modify while enumerating:");
            foreach (TeamMember member in team)
            {
                Console.WriteLine(member.ToString());
                team.AddMember(new TeamMember("Stumpy", "Lion Tamer"));
            }
        
            // Wait to continue.
            Console.WriteLine(Environment.NewLine);
            Console.WriteLine("Main method complete. Press Enter");
            Console.ReadLine();
        }
    }
}