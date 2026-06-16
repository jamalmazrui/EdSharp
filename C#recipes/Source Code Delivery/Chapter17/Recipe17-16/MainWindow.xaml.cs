using System.Windows;

namespace Apress.VisualCSharpRecipes.Chapter17
{
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();

            // Set the DataContext to a Person object
            this.DataContext = new Person()
                {
                    FirstName = "Zander",
                    LastName = "Harris"
                };
        }
    }
}
