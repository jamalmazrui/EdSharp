using System;
using System.Net;
using System.IO;
using System.Drawing;
using System.Windows.Forms;

namespace Apress.VisualCSharpRecipes.Chapter10
{
    public partial class Recipe10_04 : Form
    {
        public Recipe10_04()
        {
            InitializeComponent();
        }
 
        protected override void OnLoad(EventArgs e)
        {
            base.OnLoad(e);

            string picUri = "http://www.apress.com/img/img05/Hex_RGB4.jpg";
            string htmlUri = "http://www.apress.com";

            // Create the requests.
            WebRequest requestPic = WebRequest.Create(picUri);
            WebRequest requestHtml = WebRequest.Create(htmlUri);

            // Get the responses.
            // This takes the most significant amount of time, particularly
            // if the file is large, because the whole response is retrieved.
            WebResponse responsePic = requestPic.GetResponse();
            WebResponse responseHtml = requestHtml.GetResponse();

            // Read the image from the response stream.
            pictureBox1.Image = Image.FromStream(responsePic.GetResponseStream());

            // Read the text from the response stream.
            using (StreamReader r = 
                new StreamReader(responseHtml.GetResponseStream()))
            {
                textBox1.Text = r.ReadToEnd();
            }
        }

        [STAThread]
        public static void Main(string[] args)
        {
            Application.Run(new Recipe10_04());
        }
    }
}