using System;
using System.Drawing;
using System.Windows.Forms;

namespace Apress.VisualCSharpRecipes.Chapter08
{
    public partial class Recipe08_08 : Form
    {
        public Recipe08_08()
        {
            InitializeComponent();
        }

        Image thumbnail;

        private void Recipe08_08_Load(object sender, EventArgs e)
        {
            using (Image img = Image.FromFile("test.jpg"))
            {
                int thumbnailWidth = 0, thumbnailHeight = 0;

                // Adjust the largest dimension to 200 pixels.
                // This ensures that a thumbnail will not be larger than 
                // 200x200 pixels.
                // If you are showing multiple thumbnails, you would reserve a 
                // 200x200 pixel square for each one.
                if (img.Width > img.Height)
                {
                    thumbnailWidth = 200;
                    thumbnailHeight = Convert.ToInt32(((200F / img.Width) *
                      img.Height));
                }
                else
                {
                    thumbnailHeight = 200;
                    thumbnailWidth = Convert.ToInt32(((200F / img.Height) *
                      img.Width));
                }

                thumbnail = img.GetThumbnailImage(thumbnailWidth, thumbnailHeight,
                  null, IntPtr.Zero);
            }
        }

        private void Recipe08_08_Paint(object sender, PaintEventArgs e)
        {
            e.Graphics.DrawImage(thumbnail, 10, 10);
        }
    }
}