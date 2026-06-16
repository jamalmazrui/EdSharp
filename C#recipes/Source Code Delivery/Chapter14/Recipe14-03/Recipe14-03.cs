using System;
using System.Diagnostics;

namespace Apress.VisualCSharpRecipes.Chapter14
{
    class Recipe14_03
    {
        public static void Main () 
        {
            // If it does not exist, register an event source for this
            // application against the Application log of the local machine.
            // Trying to register an event source that already exists on the 
            // specified machine will throw a System.ArgumentException.
            if (!EventLog.SourceExists("Visual C# Recipes")) 
            {
                EventLog.CreateEventSource("Visual C# Recipes", 
                    "Application");
            }

            // Write an event to the event log.
            EventLog.WriteEntry(
                "Visual C# Recipes",      // Registered event source
                "A simple test event.",        // Event entry message
                EventLogEntryType.Information, // Event type
                1,                             // Application specific ID
                0,                             // Application specific category
                new byte[] {10, 55, 200}       // Event data
            );

            // Wait to continue.
            Console.WriteLine(Environment.NewLine);
            Console.WriteLine("Main method complete. Press Enter.");
            Console.ReadLine();
        }
    }
}
