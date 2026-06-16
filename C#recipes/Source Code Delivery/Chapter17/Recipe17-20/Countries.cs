using System.Collections.ObjectModel;

namespace Apress.VisualCSharpRecipes.Chapter17
{
    public class Countries : Collection<Country>
    {
        public Countries()
        {
            this.Add(new Country("Great Britan", Continent.Europe));
            this.Add(new Country("USA", Continent.NorthAmerica));
            this.Add(new Country("Canada", Continent.NorthAmerica));
            this.Add(new Country("France", Continent.Europe));
            this.Add(new Country("Germany", Continent.Europe));
            this.Add(new Country("Italy", Continent.Europe));
            this.Add(new Country("Spain", Continent.Europe));
            this.Add(new Country("Brazil", Continent.SouthAmerica));
            this.Add(new Country("Argentina", Continent.SouthAmerica));
            this.Add(new Country("China", Continent.Asia));
            this.Add(new Country("India", Continent.Asia));
            this.Add(new Country("Japan", Continent.Asia));
            this.Add(new Country("South Africa", Continent.Africa));
            this.Add(new Country("Tunisia", Continent.Africa));
            this.Add(new Country("Egypt", Continent.Africa));
        }
    }
}
