using System;
using System.Drawing;
using System.Windows.Forms;
using System.Drawing.Drawing2D;

namespace Apress.VisualCSharpRecipes.Chapter08
{
    public partial class Recipe08_07 : Form
    {
        public Recipe08_07()
        {
            InitializeComponent();
        }

        // Track the image size, and the type of animation
        // (expanding or shrinking).
        private bool isShrinking = false;
        private int imageSize = 0;

        // Store the logo that will be painted on the form.
        private Image image;

        private void Recipe08_07_Load(object sender, EventArgs e)
        {
            // Load the logo image from the file.
            image = Image.FromFile("test.jpg");

            // Start the timer that invalidates the form.
            tmrRefresh.Start();
        }

        private void tmrRefresh_Tick(object sender, EventArgs e)
        {
            // Change the desired image size according to the animation mode.           
            if (isShrinking)
            {
                imageSize--;
            }
            else
            {
                imageSize++;
            }

            // Change the sizing direction if it nears the form border.
            if (imageSize > (this.Width - 150))
            {
                isShrinking = true;
            }
            else if (imageSize < 1)
            {
                isShrinking = false;
            }

            // Repaint the form.
            this.Invalidate();
        }

        private void Recipe08_07_Paint(object sender, PaintEventArgs e)
        {
            Graphics g;

            g = e.Graphics;

            g.SmoothingMode = SmoothingMode.HighQuality;

            // Draw the background.
            g.FillRectangle(Brushes.Yellow, new Rectangle(new Point(0, 0),
            this.ClientSize));

            // Draw the logo image.
            g.DrawImage(image, 50, 50, 50 + imageSize, 50 + imageSize);
        }

        private void chkUseDoubleBuffering_CheckedChanged(object sender, EventArgs e)
        {
            this.DoubleBuffered = chkUseDoubleBuffering.Checked;
        }
    }
}