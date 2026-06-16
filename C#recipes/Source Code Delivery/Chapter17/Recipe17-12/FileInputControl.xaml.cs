using System.Windows.Controls;
using Microsoft.Win32;

namespace Apress.VisualCSharpRecipes.Chapter17
{
    public partial class FileInputControl : UserControl
    {
        public FileInputControl()
        {
            InitializeComponent();
        }

        private void BrowseButton_Click(
            object sender,
            System.Windows.RoutedEventArgs e)
        {
            OpenFileDialog dlg = new OpenFileDialog();
            if (dlg.ShowDialog() == true)
            {
                this.FileName = dlg.FileName;
            }
        }

        public string FileName
        {
            get
            {
                return txtBox.Text;
            }
            set
            {
                txtBox.Text = value;
            }
        }
    }
}
