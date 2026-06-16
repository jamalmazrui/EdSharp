using System;
using System.Windows.Forms;

namespace Apress.VisualCSharpRecipes.Chapter07
{
    public partial class Recipe07_20 : Form
    {
        public Recipe07_20()
        {
            // Initialization code is designer generated and contained
            // in a separate file named Recipe07-20.Designer.cs.
            InitializeComponent();
        }

        private void goButton_Click(object sender, EventArgs e)
        {
            // Navigate to the URL specified in the textbox.
            webBrowser1.Navigate(textURL.Text);
        }

        private void homeButton_Click(object sender, EventArgs e)
        {
            // Navigate to the current user's home page.
            webBrowser1.GoHome();
        }

        protected override void OnLoad(EventArgs e)
        {
            // Call the OnLoad method of the base class to ensure the Load 
            // event is raised correctly.
            base.OnLoad(e);

            // Navigate to the Apress home page when the application first
            // loads.
            webBrowser1.Navigate("http://www.apress.com");
        }

        private void backButton_Click(object sender, EventArgs e)
        {
            // Go to the previous page in the WebBrowser history.
            webBrowser1.GoBack();
        }

        private void forwarButton_Click(object sender, EventArgs e)
        {
            // Go to the next page in the WebBrowser history.
            webBrowser1.GoForward();
        }

        // Event handler to perform general interface maintenance once a document 
        // has been loaded into the WebBrowser.
        private void webBrowser1_DocumentCompleted(object sender, 
            WebBrowserDocumentCompletedEventArgs e)
        {
            // Update the content of the TextBox to reflect the current URL.
            textURL.Text = webBrowser1.Url.ToString();

            // Enable or disable the Back button depending on whether the 
            // WebBrowser has back history.
            if (webBrowser1.CanGoBack)
            {
                backButton.Enabled = true;
            }
            else
            {
                backButton.Enabled = false;
            }

            // Enable or disable the Forward button depending on whether the 
            // WebBrowser has forward history.
            if (webBrowser1.CanGoForward)
            {
                forwarButton.Enabled = true;
            }
            else
            {
                forwarButton.Enabled = false;
            }
        }

        [STAThread]
        public static void Main(string[] args)
        {
            Application.Run(new Recipe07_20());
        }
    }
}