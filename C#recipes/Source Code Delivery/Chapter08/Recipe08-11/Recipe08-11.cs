using System;
using System.Windows.Forms;
using QuartzTypeLib;

namespace Apress.VisualCSharpRecipes.Chapter08
{
    public partial class Recipe08_11 : Form
    {
        public Recipe08_11()
        {
            InitializeComponent();
        }

        private void cmdOpen_Click(object sender, EventArgs e)
        {
            // Allow the user to choose a file.
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter =
     "Media Files|*.wav;*.mp3;*.mp2;*.wma|All Files|*.*";

            if (DialogResult.OK == openFileDialog.ShowDialog())
            {
                // Access the IMediaControl interface.
                QuartzTypeLib.FilgraphManager graphManager =
                  new QuartzTypeLib.FilgraphManager();
                QuartzTypeLib.IMediaControl mc =
                  (QuartzTypeLib.IMediaControl)graphManager;

                // Specify the file.
                mc.RenderFile(openFileDialog.FileName);

                // Start playing the audio asynchronously.
                mc.Run();
            }
        }
    }
}