using System;
using System.Threading;
using System.Windows.Forms;

namespace Apress.VisualCSharpRecipes.Chapter07
{
    public partial class Recipe07_08 : Form
    {

        public Recipe07_08()
        {
            // Initialization code is designer generated and contained
            // in a separate file named Recipe07-08.Designer.cs.
            InitializeComponent();
        }

        private void btnTime_Click(object sender, EventArgs e)
        {
            // Set the input mask to that of a short time.
            this.mskTextBox.UseSystemPasswordChar = false;
            this.mskTextBox.Mask = "00:00";
            this.lblActiveMask.Text = this.mskTextBox.Mask;
            this.mskTextBox.Focus();
        }

        private void btnUSZip_Click(object sender, EventArgs e)
        {
            // Set the input mask to that of a US ZIP code.
            this.mskTextBox.UseSystemPasswordChar = false;
            this.mskTextBox.Mask = "00000-9999";
            this.lblActiveMask.Text = this.mskTextBox.Mask;
            this.mskTextBox.Focus();
        }

        private void btnUKPost_Click(object sender, EventArgs e)
        {
            // Set the input mask to that of a UK postcode.
            this.mskTextBox.UseSystemPasswordChar = false;
            this.mskTextBox.Mask = ">LCCC 9LL";
            this.lblActiveMask.Text = this.mskTextBox.Mask;
            this.mskTextBox.Focus();
        }

        private void btnCurrency_Click(object sender, EventArgs e)
        {
            // Set the input mask to that of a currency.
            this.mskTextBox.UseSystemPasswordChar = false;
            this.mskTextBox.Mask = "$999,999.00";
            this.lblActiveMask.Text = this.mskTextBox.Mask;
            this.mskTextBox.Focus();
        }

        private void btnDate_Click(object sender, EventArgs e)
        {
            // Set the input mask to that of a short date.
            this.mskTextBox.UseSystemPasswordChar = false;
            this.mskTextBox.Mask = "00/00/0000";
            this.lblActiveMask.Text = this.mskTextBox.Mask;
            this.mskTextBox.Focus();
        }

        private void btnSecret_Click(object sender, EventArgs e)
        {
            // Set the input mask to that of a secret PIN.
            this.mskTextBox.UseSystemPasswordChar = true;
            this.mskTextBox.Mask = "0000";
            this.lblActiveMask.Text = this.mskTextBox.Mask;
            this.mskTextBox.Focus();
        }

        [STAThread]
        public static void Main(string[] args)
        {
            Application.Run(new Recipe07_08());
        }
    }
}