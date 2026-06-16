using System;
using System.Drawing;
using System.Windows.Forms;

namespace Apress.VisualCSharpRecipes.Chapter08
{
    public partial class Recipe08_05 : Form
    {
        public Recipe08_05()
        {
            InitializeComponent();
        }

        private void Recipe08_05_Load(object sender, EventArgs e)
        {
            string text = "The quick brown fox jumps over the lazy dog.";
            using (Font font = new Font("Tahoma", 20))
            {
                // Create an in-memory bitmap.
                Bitmap b = new Bitmap(600, 600);
                using (Graphics g = Graphics.FromImage(b))
                {
                    g.FillRectangle(Brushes.White, new Rectangle(0, 0, 
                        b.Width, b.Height));

                    // Draw several lines of text on the bitmap.
                    for (int i = 0; i < 10; i++)
                    {
                        g.DrawString(text, font, Brushes.Black, 
                            50, 50 + i * 60);
                    }
                }

                // Display the bitmap in the picture box.
                pictureBox1.BackgroundImage = b;
                pictureBox1.Size = b.Size;
            }
        }
    }
}