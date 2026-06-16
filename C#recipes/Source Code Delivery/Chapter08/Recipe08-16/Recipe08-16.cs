using System;
using System.Drawing;
using System.Windows.Forms;
using System.Drawing.Printing;

namespace Apress.VisualCSharpRecipes.Chapter08
{
    public partial class Recipe08_16 : Form
    {
        public Recipe08_16()
        {
            InitializeComponent();
        }

        private void cmdPrint_Click(object sender, EventArgs e)
        {
            // Create the document and attach an event handler.
            string text = "Windows Server 2003 builds on the core strengths " +
              "of the Windows family of operating systems--security, " +
              "manageability, reliability, availability, and scalability. " +
              "Windows Server 2003 provides an application environment to " +
              "build, deploy, manage, and run XML Web services. " +
              "Additionally, advances in Windows Server 2003 provide many " +
              "benefits for developing applications.";
            PrintDocument doc = new ParagraphDocument(text);
            doc.PrintPage += new PrintPageEventHandler(this.Doc_PrintPage);

            // Allow the user to choose a printer and specify other settings.
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
            ParagraphDocument doc = (ParagraphDocument)sender;

            // Define the font and text.
            using (Font font = new Font("Arial", 15))
            {
                e.Graphics.DrawString(doc.Text, font, Brushes.Black,
                   e.MarginBounds, StringFormat.GenericDefault);
            }
        }
    }

    public class ParagraphDocument : PrintDocument
    {
        private string text;
        public string Text
        {
            get { return text; }
            set { text = value; }
        }

        public ParagraphDocument(string text)
        {
            this.Text = text;
        }
    }
}