using System;
using System.Windows.Forms;

namespace Apress.VisualCSharpRecipes.Chapter07
{
    public partial class Recipe07_04 : Form
    {
        public Recipe07_04()
        {
            // Initialization code is designer generated and contained
            // in a separate file named Recipe07-04.Designer.cs.
            InitializeComponent();
        }

        // Override the OnLoad method to show the initial list of forms.
        protected override void OnLoad(EventArgs e)
        {
            // Call the OnLoad method of the base class to ensure the Load 
            // event is raised correctly.
            base.OnLoad(e);

            // Refresh the list to display the initial set of forms.
            this.RefreshForms();
        }

        // A button click event handler to create a new child form.
        private void btnNewForm_Click(object sender, EventArgs e)
        {
            // Create a new child form and set its name as specified.
            // If no name is specified, use a default name.
            Recipe07_04Child child = new Recipe07_04Child();

            if (String.IsNullOrEmpty(this.txtFormName.Text))
            {
                child.Name = "Child Form";
            }
            else
            {
                child.Name = this.txtFormName.Text;
            }
            
            // Show the new child form.
            child.Show();
        }

        // List selection event handler to activate the selected form based on 
        // its name.
        private void listForms_SelectedIndexChanged(object sender, EventArgs e)
        {
            // Activate the selected form using its name as the index into the
            // collection of active forms. If there are duplicate forms with the
            // same name, the first one found will be activated.
            Form form = Application.OpenForms[this.listForms.Text];

            // If the form has been closed, using its name as an index into the 
            // FormCollection will return null. In this instance, update the 
            // list of forms.
            if (form != null)
            {
                // Activate the selected form.
                form.Activate();
            }
            else
            {
                // Display a message and refresh the form list.
                MessageBox.Show("Form closed; refreshing list...",
                    "Form Closed");
                this.RefreshForms();
            }
        }

        // A button click event handler to initiate a refresh of the list of 
        // active forms.
        private void btnRefresh_Click(object sender, EventArgs e)
        {
            RefreshForms();
        }

        // A method to perform a refresh of the list of active forms.
        private void RefreshForms()
        {
            // Clear the list and repopulate from the Application.OpenForms
            // property.
            this.listForms.Items.Clear();

            foreach (Form f in Application.OpenForms)
            {
                this.listForms.Items.Add(f.Name);
            }
        }

        [STAThread]
        public static void Main(string[] args)
        {
            Application.Run(new Recipe07_04());
        }
    }
}