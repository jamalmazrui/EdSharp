using System.Windows;

namespace Apress.VisualCSharpRecipes.Chapter17
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        // Handles Clear Button click event.
        private void btnClear_Click(object sender, RoutedEventArgs e)
        {
            // Select all the text in the FlowDocument and cut it.
            rtbTextBox1.SelectAll();
            rtbTextBox1.Cut();
        }
    }
}
