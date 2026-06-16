using System;
using System.Windows.Forms;
using QuartzTypeLib;

namespace Apress.VisualCSharpRecipes.Chapter08
{
    public partial class Recipe08_12 : Form
    {
        public Recipe08_12()
        {
            InitializeComponent();
        }

        // Define constants used for specifying the window style.
        private const int WS_CHILD = 0x40000000;
        private const int WS_CLIPCHILDREN = 0x2000000;

        // Hold a form-level reference to the media control interface,
        // so the code can control playback of the currently loaded
        // movie.
        private IMediaControl mc = null;

        // Hold a form-level reference to the video window in case it
        // needs to be resized.
        private IVideoWindow videoWindow = null;

        private void cmdOpen_Click(object sender, EventArgs e)
        {
            // Allow the user to choose a file.
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter =
            "Media Files|*.mpg;*.avi;*.wma;*.mov;*.wav;*.mp2;*.mp3|" + 
            "All Files|*.*";

            if (DialogResult.OK == openFileDialog.ShowDialog())
            {
                // Stop the playback for the current movie, if it exists.
                if (mc != null) mc.Stop();

                // Load the movie file.
                FilgraphManager graphManager = new FilgraphManager();
                graphManager.RenderFile(openFileDialog.FileName);

                // Attach the view to a picture box on the form.
                try
                {
                    videoWindow = (IVideoWindow)graphManager;
                    videoWindow.Owner = (int)pictureBox1.Handle;
                    videoWindow.WindowStyle = WS_CHILD | WS_CLIPCHILDREN;
                    videoWindow.SetWindowPosition(
                      pictureBox1.ClientRectangle.Left,
                      pictureBox1.ClientRectangle.Top,
                      pictureBox1.ClientRectangle.Width,
                      pictureBox1.ClientRectangle.Height);
                }
                catch
                {
                    // An error can occur if the file does not have a video
                    // source (for example, an MP3 file.)
                    // You can ignore this error and still allow playback to
                    // continue (without any visualization).
                }

                // Start the playback (asynchronously).
                mc = (IMediaControl)graphManager;
                mc.Run();
            }
        }

        private void pictureBox1_SizeChanged(object sender, EventArgs e)
        {
            if (videoWindow != null)
            {
                try
                {
                    videoWindow.SetWindowPosition(
                        pictureBox1.ClientRectangle.Left,
                        pictureBox1.ClientRectangle.Top,
                        pictureBox1.ClientRectangle.Width,
                        pictureBox1.ClientRectangle.Height);
                }
                catch
                {
                    // Ignore the exception thrown when resizing the form
                    // when the file does not have a video source.
                }
            }
        }
    }
}