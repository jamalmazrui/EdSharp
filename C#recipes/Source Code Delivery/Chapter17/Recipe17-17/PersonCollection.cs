using System.Collections.ObjectModel;

namespace Apress.VisualCSharpRecipes.Chapter17
{
    public class PersonCollection : Collection<Person>
    {
        public PersonCollection()
        {
            this.Add(new Person()
                         {
                             FirstName = "Sam",
                             LastName = "Bourton",
                             Age = 33,
                             Photo = "samb.png"
                         });
            this.Add(new Person()
                         {
                             FirstName = "Sam",
                             LastName = "Noble",
                             Age = 24,
                             Photo = "samn.jpg"
                         });
        }
    }
}