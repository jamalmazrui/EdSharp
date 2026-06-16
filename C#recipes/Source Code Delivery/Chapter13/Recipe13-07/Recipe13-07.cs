using System;

namespace Apress.VisualCSharpRecipes.Chapter13
{
    public class Person : IFormattable
    {
        // Private members to hold the person's title and name details.
        private string title;
        private string[] names;

        // Constructor used to set the person's title and names.
        public Person(string title, params string[] names)
        {
            this.title = title;
            this.names = names;
        }

        // Override the Object.ToString method to return the person's
        // name using the general format.
        public override string ToString()
        {
            return ToString("G", null);
        }

        // Implementation of the IFormattable.ToString method to return the 
        // person's name in different forms based on the format string
        // provided.
        public string ToString(string format, IFormatProvider formatProvider)
        {
            string result = null;

            // Use the general format if none is specified.
            if (format == null) format = "G";

            // The contents of the format string determine the format of the
            // name returned.
            switch (format.ToUpper()[0])
            {
                case 'S':
                    // Use short form - first initial and surname.
                    result = names[0][0] + ". " + names[names.Length - 1];
                    break;

                case 'P':
                    // Use polite form - title, initials, and surname
                    // Add the person's title to the result.
                    if (title != null && title.Length != 0)
                    {
                        result = title + ". ";
                    }
                    // Add the person's initials and surname.
                    for (int count = 0; count < names.Length; count++)
                    {
                        if (count != (names.Length - 1))
                        {
                            result += names[count][0] + ". ";
                        }
                        else
                        {
                            result += names[count];
                        }
                    }
                    break;

                case 'I':
                    // Use informal form - first name only.
                    result = names[0];
                    break;

                case 'G':
                default:
                    // Use general/default form - first name and surname.
                    result = names[0] + " " + names[names.Length - 1];
                    break;
            }
            return result;
        }
    }

    // A class to demonstrate the use of Person.
    public class Recipe13_07
    {
        public static void Main()
        {
            // Create a Person object representing a man with the name
            // Mr. Richard Glen David Peters.
            Person person =
                new Person("Mr", "Richard", "Glen", "David", "Peters");

            // Display the person's name using a variety of format strings.
            System.Console.WriteLine("Dear {0:G},", person);
            System.Console.WriteLine("Dear {0:P},", person);
            System.Console.WriteLine("Dear {0:I},", person);
            System.Console.WriteLine("Dear {0},", person);
            System.Console.WriteLine("Dear {0:S},", person);

            // Wait to continue.
            Console.WriteLine(Environment.NewLine); 
            Console.WriteLine("Main method complete. Press Enter");
            Console.ReadLine();
        }
    }
}