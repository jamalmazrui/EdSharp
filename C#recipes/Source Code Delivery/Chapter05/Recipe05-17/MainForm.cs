using System;
using System.Windows.Forms;

namespace Apress.VisualCSharpRecipes.Chapter05
{
    public partial class MainForm : Form
    {
        public MainForm()
        {
            InitializeComponent();
        }

        private void mnuOpen_Click(object sender, EventArgs e)
        {
            OpenFileDialog dlg = new OpenFileDialog();
            dlg.Filter = "Rich Text Files (*.rtf)|*.RTF|" +
              "All files (*.*)|*.*";
            dlg.CheckFileExists = true;
            dlg.InitialDirectory = Application.StartupPath;

            if (dlg.ShowDialog() == DialogResult.OK)
            {
                rtDoc.LoadFile(dlg.FileName);
                rtDoc.Enabled = true;
            }
        }

        private void mnuSave_Click(object sender, EventArgs e)
        {
            SaveFileDialog dlg = new SaveFileDialog();
            dlg.Filter = "RichText Files (*.rtf)|*.RTF|Text Files (*.txt)|*.TXT" +
              "|All files (*.*)|*.*";
            dlg.CheckFileExists = true;
            dlg.InitialDirectory = Application.StartupPath;

            if (dlg.ShowDialog() == DialogResult.OK)
            {
                rtDoc.SaveFile(dlg.FileName);
            }
        }

        private void mnuExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}