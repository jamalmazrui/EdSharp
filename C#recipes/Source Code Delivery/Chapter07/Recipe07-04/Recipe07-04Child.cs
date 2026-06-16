using System;
using System.Windows.Forms;

namespace Apress.VisualCSharpRecipes.Chapter07
{
    public partial class Recipe07_04Child : Form
    {
        public Recipe07_04Child()
        {
            // Initialization code is designer generated and contained
            // in a separate file named Recipe07-04Child.Designer.cs.
            InitializeComponent();
        }

        // Override the OnPaint method to correctly display the name of the 
        // form.
        protected override void OnPaint(PaintEventArgs e)
        {
            // Call the OnPaint method of the base class to ensure the Paint 
            // event is raised correctly.
            base.OnPaint(e);

            // Display the name of the form.
            this.lblFormName.Text = this.Name;
        }

        // A button click event handler to close the child form.
        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}