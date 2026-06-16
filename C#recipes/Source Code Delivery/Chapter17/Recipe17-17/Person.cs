using System.Collections.ObjectModel;

namespace Apress.VisualCSharpRecipes.Chapter17
{
    public class Person
    {
        public string FirstName
        {
            get;
            set;
        }
        public string LastName
        {
            get;
            set;
        }
        public int Age
        {
            get;
            set;
        }
        public string Photo
        {
            get;
            set;
        }

        public override string ToString()
        {
            return string.Format("{0} {1}", FirstName, LastName);
        }
    }
}