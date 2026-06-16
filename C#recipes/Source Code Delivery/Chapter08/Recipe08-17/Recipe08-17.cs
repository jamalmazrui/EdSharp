using System;
using System.Drawing;
using System.Windows.Forms;
using System.Drawing.Printing;

namespace Apress.VisualCSharpRecipes.Chapter08
{
    public partial class Recipe08_17 : Form
    {
        public Recipe08_17()
        {
            InitializeComponent();
        }

        private PrintDocument doc;

        private void Recipe08_17_Load(object sender, EventArgs e)
        {
            // Set the allowed zoom settings.
            for (int i = 1; i <= 10; i++)
            {
                lstZoom.Items.Add((i * 10).ToString());
            }

            // Create a document with 100 lines.
            string[] printText = new string[100];
            for (int i = 0; i < 100; i++)
            {
                printText[i] = i.ToString();
                printText[i] += ": The quick brown fox jumps over the lazy dog.";
            }

            doc = new TextDocument(printText);
            doc.PrintPage += this.Doc_PrintPage;

            lstZoom.Text = "100";
            printPreviewControl.Zoom = 1;
            printPreviewControl.Document = doc;
            printPreviewControl.Rows = 2;
        }

        private void cmdPrint_Click(object sender, EventArgs e)
        {
            // Set the zoom.
            printPreviewControl.Zoom = Single.Parse(lstZoom.Text) / 100;

            // Show the full two pages, one above the other.
            printPreviewControl.Rows = 2;

            // Rebind the PrintDocument to refresh the preview.
            printPreviewControl.Document = doc;
        }

        private void Doc_PrintPage(object sender, PrintPageEventArgs e)
        {
            // Retrieve the document that sent this event.
            TextDocument doc = (TextDocument)sender;

            // Define the font and determine the line height.
            using (Font font = new Font("Arial", 10))
            {
                float lineHeight = font.GetHeight(e.Graphics);

                // Create variables to hold position on page.
                float x = e.MarginBounds.Left;
                float y = e.MarginBounds.Top;

                // Increment the page counter (to reflect the page that 
                // is about to be printed).
                doc.PageNumber += 1;

                // Print all the information that can fit on the page.        
                // This loop ends when the next line would go over the
                // margin bounds, or there are no more lines to print.

                while ((y + lineHeight) < e.MarginBounds.Bottom &&
                  doc.Offset <= doc.Text.GetUpperBound(0))
                {
                    e.Graphics.DrawString(doc.Text[doc.Offset], font,
                      Brushes.Black, x, y);

                    // Move to the next line of data.
                    doc.Offset += 1;

                    // Move the equivalent of one line down the page.
                    y += lineHeight;
                }

                if (doc.Offset < doc.Text.GetUpperBound(0))
                {
                    // There is still at least one more page.
                    // Signal this event to fire again.
                    e.HasMorePages = true;
                }
                else
                {
                    // Printing is complete.
                    doc.Offset = 0;
                }
            }
        }
    }

    class TextDocument : PrintDocument
    {
        private string[] text;
        private int pageNumber;
        private int offset;

        public string[] Text
        {
            get { return text; }
            set { text = value; }
        }

        public int PageNumber
        {
            get { return pageNumber; }
            set { pageNumber = value; }
        }

        public int Offset
        {
            get { return offset; }
            set { offset = value; }
        }

        public TextDocument(string[] text)
        {
            this.Text = text;
        }
    }
}