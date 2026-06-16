using System.Collections.ObjectModel;

namespace Apress.VisualCSharpRecipes.Chapter17
{
    public class PersonCollection : ObservableCollection<Person>
    {
        public PersonCollection()
        {
            this.Add(new Person()
                         {
                             FirstName = "Sam",
                             LastName = "Bourton",
                             Age = 33,
                             Occupation = "Engineer"
                         });
            this.Add(new Person()
                         {
                             FirstName = "Adam",
                             LastName = "Freeman",
                             Age = 37,
                             Occupation = "Professional"
                         });
            this.Add(new Person()
                         {
                             FirstName = "Sam",
                             LastName = "Noble",
                             Age = 24,
                             Occupation = "Engineer"
                         });
        }
    }
}