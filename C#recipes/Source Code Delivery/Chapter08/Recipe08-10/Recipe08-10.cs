using System;
using System.Windows.Forms;
using System.Media;

namespace Apress.VisualCSharpRecipes.Chapter08
{
    public partial class Recipe08_10 : Form
    {
        public Recipe08_10()
        {
            InitializeComponent();
        }

        private void cmdOpen_Click(object sender, EventArgs e)
        {
            // Allow the user to choose a file.
            OpenFileDialog openDialog = new OpenFileDialog();
            openDialog.InitialDirectory = "C:\\Windows\\Media ";
            openDialog.Filter = "WAV Files|*.wav|All Files|*.*";

            if (DialogResult.OK == openDialog.ShowDialog())
            {
                SoundPlayer player = new SoundPlayer(openDialog.FileName);

                try
                {
                    player.Play();
                }
                catch (Exception)
                {
                    MessageBox.Show("An error occurred while playing media.");
                }
                finally
                {
                    player.Dispose();
                }
            }
        }
    }
}