using System;

namespace Apress.VisualCSharpRecipes.Chapter13
{
    [Serializable]
    public sealed class MailReceivedEventArgs : EventArgs
    {
        // Private read-only members that hold the event state that is to be
        // distributed to all event handlers. The MailReceivedEventArgs class
        // will specify who sent the received mail and what the subject is.
        private readonly string from;
        private readonly string subject;

        // Constructor, initializes event state.
        public MailReceivedEventArgs(string from, string subject)
        {
            this.from = from;
            this.subject = subject;
        }

        // Read-only properties to provide access to event state.
        public string From { get { return from; } }
        public string Subject { get { return subject; } }
    }

    // A class to demonstrate the use of MailReceivedEventArgs.
    public class Recipe13_09
    {
        public static void Main()
        {
            MailReceivedEventArgs args = 
                new MailReceivedEventArgs("Danielle", "Your book");

            Console.WriteLine("From: {0}, Subject: {1}", args.From, args.Subject);

            // Wait to continue.
            Console.WriteLine(Environment.NewLine);
            Console.WriteLine("Main method complete. Press Enter");
            Console.ReadLine();
        }
    }
}