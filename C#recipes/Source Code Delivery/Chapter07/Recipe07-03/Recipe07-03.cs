using System;
using System.Windows.Forms;

namespace Apress.VisualCSharpRecipes.Chapter07
{
    public partial class Recipe07_03 : Form
    {
        public Recipe07_03()
        {
            // Initialization code is designer generated and contained
            // in a separate file named Recipe07-03.Designer.cs.
            InitializeComponent();
        }

        // The event handler for the button click event.
        private void cmdProcessAll_Click(object sender, System.EventArgs e)
        {
            ProcessControls(this);
        }

        private void ProcessControls(Control ctrl)
        {
            // Ignore the control unless it's a textbox.
            if (ctrl is TextBox)
            {
                ctrl.Text = "";
            }

            // Process controls recursively. 
            // This is required if controls contain other controls
            // (for example, if you use panels, group boxes, or other
            // container controls).
            foreach (Control ctrlChild in ctrl.Controls)
            {
                ProcessControls(ctrlChild);
            }
        }

        [STAThread]
        public static void Main(string[] args)
        {
            Application.Run(new Recipe07_03());
        }
    }
}