using System;
using System.Drawing;
using System.Windows.Forms;
using System.Drawing.Printing;

namespace Apress.VisualCSharpRecipes.Chapter08
{
    public partial class Recipe08_15 : Form
    {
        public Recipe08_15()
        {
            InitializeComponent();
        }

        private void cmdPrint_Click(object sender, EventArgs e)
        {
            // Create a document with 100 lines.
            string[] printText = new string[101];
            for (int i = 0; i < 101; i++)
            {
                printText[i] = i.ToString();
                printText[i] += 
                    ": The quick brown fox jumps over the lazy dog.";
            }

            PrintDocument doc = new TextDocument(printText);
            doc.PrintPage += this.Doc_PrintPage;

            PrintDialog dlgSettings = new PrintDialog();
            dlgSettings.Document = doc;

            // If the user clicked OK, print the document.
            if (dlgSettings.ShowDialog() == DialogResult.OK)
            {
                doc.Print();
            }
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