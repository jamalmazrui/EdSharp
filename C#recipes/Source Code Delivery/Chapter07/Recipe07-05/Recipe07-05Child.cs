using System;
using System.Windows.Forms;

namespace Apress.VisualCSharpRecipes.Chapter07
{
    // An MDI child form.
    public partial class Recipe07_05Child : Form
    {
        public Recipe07_05Child()
        {
            // Initialization code is designer generated and contained
            // in a separate file named Recipe07-05Child.Designer.cs.
            InitializeComponent();
        }

        // When a button on any of the MDI child forms is clicked, display the 
        // contents of a each form by enumerating the MdiChildren collection.
        private void cmdShowAllWindows_Click(object sender, EventArgs e)
        {
            foreach (Form frm in this.MdiParent.MdiChildren)
            {
                // Cast the generic Form to the Recipe07_05Childderived class 
                // type.
                Recipe07_05Child child = (Recipe07_05Child)frm;
                MessageBox.Show(child.LabelText, frm.Text);
            }
        }

        // On load, set the MDI Child form's label to the current date/time.
        protected override void OnLoad(EventArgs e)
        {
            // Call the OnLoad method of the base class to ensure the Load
            // event is raised correctly.
            base.OnLoad(e);

            label.Text = DateTime.Now.ToString();
        }

        // A property to provide easy access to the label data.
        public string LabelText
        {
            get { return label.Text; }
        }
    }
}