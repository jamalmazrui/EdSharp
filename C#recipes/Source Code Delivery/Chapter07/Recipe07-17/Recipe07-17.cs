using System;
using System.Windows.Forms;
using System.Text.RegularExpressions;

namespace Apress.VisualCSharpRecipes.Chapter07
{
    public partial class Recipe07_17 : Form
    {
        public Recipe07_17()
        {
            // Initialization code is designer generated and contained
            // in a separate file named Recipe07-17.Designer.cs.
            InitializeComponent();
        }

        // Button click event handler ensures the ErrorProvider is not
        // reporting any error for each control before proceeding.
        private void Button1_Click(object sender, EventArgs e)
        {
            string errorText = "";
            bool invalidInput = false;

            foreach (Control ctrl in this.Controls)
            {
                if (errProvider.GetError(ctrl) != "")
                {
                    errorText += "   * " + errProvider.GetError(ctrl) + "\n";
                    invalidInput = true;
                }
            }

            if (invalidInput)
            {   
                MessageBox.Show(
                    "The form contains the following unresolved errors:\n\n" +
                    errorText, "Invalid Input", MessageBoxButtons.OK, 
                    MessageBoxIcon.Warning);
            }
            else
            {
                this.Close();
            }
        }

        // When the TextBox loses focus, check that the contents are a valid
        // e-mail address.
        private void txtEmail_Leave(object sender, EventArgs e)
        {
            // Create a regular expression to check for valid e-mail addresses.
            Regex regex = new Regex(@"^[\w-]+@([\w-]+\.)+[\w-]+$");

            // Validate the text from the control that raised the event.
            Control ctrl = (Control)sender;
            if (String.IsNullOrEmpty(ctrl.Text) || regex.IsMatch(ctrl.Text))
            {
                errProvider.SetError(ctrl, "");
            }
            else
            {
                errProvider.SetError(ctrl, "This is not a valid email address.");
            }
        }

        [STAThread]
        public static void Main(string[] args)
        {
            Application.Run(new Recipe07_17());
        }
    }
}