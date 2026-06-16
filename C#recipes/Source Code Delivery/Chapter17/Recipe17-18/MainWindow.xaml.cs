using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace Apress.VisualCSharpRecipes.Chapter17
{
    public partial class MainWindow : Window
    {
        // Create an instance of the PersonCollection class
        PersonCollection people =
            new PersonCollection();

        public MainWindow()
        {
            InitializeComponent();

            // Set the DataContext to the PersonCollection
            this.DataContext = people;
        }

        private void AddButton_Click(
            object sender, RoutedEventArgs e)
        {
            people.Add(new Person()
            {
                FirstName = "Simon",
                LastName = "Williams",
                Age = 39,
                Occupation = "Professional"
            });
        }
    }
}
