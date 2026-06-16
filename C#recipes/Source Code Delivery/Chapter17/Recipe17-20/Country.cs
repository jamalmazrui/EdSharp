using System.Collections.ObjectModel;

namespace Apress.VisualCSharpRecipes.Chapter17
{
    public class Country
    {
        public Country(string name, Continent continent)
        {
            Name = name;
            Continent = continent;
        }

        public string Name { get; set; }
        public Continent Continent  { get; set; }
    }
}
