using System;

namespace Apress.VisualCSharpRecipes.Chapter03
{
    [AttributeUsage(AttributeTargets.Class | AttributeTargets.Assembly,
         AllowMultiple = true, Inherited = false)]
    public class AuthorAttribute : System.Attribute
    {
        private string company; // Creator's company
        private string name;    // Creator's name

        // Declare a public constructor.
        public AuthorAttribute(string name)
        {
            this.name = name;
            company = "";
        }

        // Declare a property to get/set the company field.
        public string Company
        {
            get { return company; }
            set { company = value; }
        }

        // Declare a property to get the internal field.
        public string Name
        {
            get { return name; }
        }
    }
}
