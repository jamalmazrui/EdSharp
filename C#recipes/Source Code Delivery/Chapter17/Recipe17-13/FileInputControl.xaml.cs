using System.Diagnostics;
using System.IO;
using System.Windows.Controls;
using System.Windows.Input;
using Microsoft.Win32;

namespace Apress.VisualCSharpRecipes.Chapter17
{
    public partial class FileInputControl : UserControl
    {
        public FileInputControl()
        {
            InitializeComponent();

            // Register command bindings

            // ApplicationCommands.Find
            CommandManager.RegisterClassCommandBinding(
                typeof(FileInputControl),
                new CommandBinding(
                    ApplicationCommands.Find,
                    FindCommand_Executed,
                    FindCommand_CanExecute));

            // ApplicationCommands.Open
            CommandManager.RegisterClassCommandBinding(
                typeof(FileInputControl),
                new CommandBinding(
                    ApplicationCommands.Open,
                    OpenCommand_Executed,
                    OpenCommand_CanExecute));
        }

        #region Find Command

        private void FindCommand_CanExecute(
            object sender,
            CanExecuteRoutedEventArgs e)
        {
            e.CanExecute = true;
        }

        private void FindCommand_Executed(
            object sender,
            ExecutedRoutedEventArgs e)
        {
            DoFindFile();
        }

        #endregion

        #region Open Command

        private void OpenCommand_CanExecute(
            object sender,
            CanExecuteRoutedEventArgs e)
        {
            e.CanExecute =
                !string.IsNullOrEmpty(this.FileName)
                && File.Exists(this.FileName);
        }

        private void OpenCommand_Executed(
            object sender,
            ExecutedRoutedEventArgs e)
        {
            Process.Start(this.FileName);
        }

        #endregion

        private void BrowseButton_Click(
            object sender,
            System.Windows.RoutedEventArgs e)
        {
            DoFindFile();
        }

        private void DoFindFile()
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
