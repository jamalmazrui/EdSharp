using System;
using System.Drawing;
using System.Windows.Forms;
using System.Drawing.Printing;

namespace Apress.VisualCSharpRecipes.Chapter08
{
    public partial class Recipe08_14 : Form
    {
        public Recipe08_14()
        {
            InitializeComponent();
        }

        private void cmdPrint_Click(object sender, EventArgs e)
        {

            // Create the document and attach an event handler.
            PrintDocument doc = new PrintDocument();
            doc.PrintPage += this.Doc_PrintPage;

            // Allow the user to choose a printer and specify other settings.
            PrintDialog dlgSettings = new PrintDialog();
            dlgSettings.Document = doc;

            // enable the following line for 64-bit windows
            //dlgSettings.UseEXDialog = true;


            // If the user clicked OK, print the document.
            if (dlgSettings.ShowDialog() == DialogResult.OK)
            {
                // This method returns immediately, before the print job starts.
                // The PrintPage event will fire asynchronously.
                doc.Print();
            }

        }

        private void Doc_PrintPage(object sender, PrintPageEventArgs e)
        {
            // Define the font.
            using (Font font = new Font("Arial", 30))
            {
                // Determine the position on the page.
                // In this case, we read the margin settings
                // (although there is nothing that prevents your code
                // from going outside the margin bounds.)
                float x = e.MarginBounds.Left;
                float y = e.MarginBounds.Top;

                // Determine the height of a line (based on the font used).
                float lineHeight = font.GetHeight(e.Graphics);

                // Print five lines of text.
                for (int i = 0; i < 5; i++)
                {

                    // Draw the text with a black brush,
                    // using the font and coordinates we have determined.
                    e.Graphics.DrawString("This is line " + i.ToString(),
                      font, Brushes.Black, x, y);

                    // Move down the equivalent spacing of one line.
                    y += lineHeight;
                }
                y += lineHeight;

                // Draw an image.
                e.Graphics.DrawImage(Image.FromFile(Application.StartupPath +
                  "\\test.jpg"), x, y);
            }
        }
    }
}