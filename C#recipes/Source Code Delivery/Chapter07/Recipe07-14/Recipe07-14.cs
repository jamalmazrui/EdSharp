using System;
using System.Windows.Forms;

namespace Apress.VisualCSharpRecipes.Chapter07
{
    public partial class Recipe07_14 : Form
    {
        public Recipe07_14()
        {
            // Initialization code is designer generated and contained
            // in a separate file named Recipe07-17.cs.
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Form form = new Form();
            form.FormBorderStyle = FormBorderStyle.None;
            form.Show();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Form form = new Form();
            form.ControlBox = false;
            form.MinimizeBox = false;
            form.MaximizeBox = false;
            form.FormBorderStyle = FormBorderStyle.FixedSingle;
            form.Text = String.Empty;
            form.Show();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        [STAThread]
        public static void Main(string[] args)
        {
            Application.Run(new Recipe07_14());
        }
    }
}
