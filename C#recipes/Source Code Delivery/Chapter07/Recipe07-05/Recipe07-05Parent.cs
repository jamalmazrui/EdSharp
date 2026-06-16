using System;
using System.Windows.Forms;

namespace Apress.VisualCSharpRecipes.Chapter07
{
    // An MDI parent form.
    public partial class Recipe07_05Parent : Form
    {
        public Recipe07_05Parent()
        {
            // Initialization code is designer generated and contained
            // in a separate file named Recipe07-05Parent.Designer.cs.
            InitializeComponent();
        }

        // When the New menu item is clicked, create a new MDI child.
        private void mnuNew_Click(object sender, EventArgs e)
        {
            Recipe07_05Child frm = new Recipe07_05Child();
            frm.MdiParent = this;
            frm.Show();
        }

        [STAThread]
        public static void Main(string[] args)
        {
            Application.Run(new Recipe07_05Parent());
        }
    }
}