using System;
using System.ComponentModel;
using System.Windows.Forms;
using Apress.VisualCSharpRecipes.Chapter07.Properties;

namespace Apress.VisualCSharpRecipes.Chapter07
{
    public partial class Recipe07_06 : Form
    {
        public Recipe07_06()
        {
            // Initialization code is designer generated and contained
            // in a separate file named Recipe07-06.Designer.cs.
            InitializeComponent();
        }

        private void Button_Click(object sender, EventArgs e)
        {
            // Change the color of the textbox depending on which button
            // was pressed.
            Button btn = sender as Button;

            if (btn != null)
            {
                // Set the background color of the textbox.
                textBox1.BackColor = btn.ForeColor;

                // Update the application settings with the new value.
                Settings.Default.Color = textBox1.BackColor;
            }
        }

        protected override void OnClosing(CancelEventArgs e)
        {
            // Call the OnClosing method of the base class to ensure the 
            // FormClosing event is raised correctly.
            base.OnClosing(e);

            // Update the application settings for Form.
            Settings.Default.Size = this.Size;

            // Store all application settings.
            Settings.Default.Save();
        }

        [STAThread]
        public static void Main(string[] args)
        {
            Application.Run(new Recipe07_06());
        }
    }
}