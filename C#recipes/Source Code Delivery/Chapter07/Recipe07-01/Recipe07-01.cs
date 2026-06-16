using System;
using System.Windows.Forms;

namespace Apress.VisualCSharpRecipes.Chapter07
{
    public partial class Recipe07_01 : Form
    {
        public Recipe07_01()
        {
            // Initialization code is designer generated and contained
            // in a separate file named Recipe07-01.Designer.cs.
            InitializeComponent();
        }

        protected override void OnLoad(EventArgs e)
        {
            // Call the OnLoad method of the base class to ensure the Load 
            // event is raised correctly.
            base.OnLoad(e);

            // Create an array of strings to use as the labels for
            // the dynamic checkboxes.
            string[] foods = {"Grain", "Bread", "Beans", "Eggs",
                                "Chicken", "Milk", "Fruit", "Vegetables",
                                "Pasta", "Rice", "Fish", "Beef"};

            // Suspend the form's layout logic while multiple controls 
            // are added.
            this.SuspendLayout();

            // Specify the Y coordinate of the topmost checkbox in the list.
            int topPosition = 10;

            // Create one new checkbox for each name in the list of 
            // food types.
            foreach (string food in foods)
            {
                // Create a new checkbox.
                CheckBox checkBox = new CheckBox();

                // Configure the new checkbox.
                checkBox.Top = topPosition;
                checkBox.Left = 10;
                checkBox.Text = food;

                // Set the Y coordinate of the next checkbox.
                topPosition += 30;

                // Add the checkbox to the panel contained by the form.
                panel1.Controls.Add(checkBox);
            }

            // Resume the forms layout logic now that all controls
            // have been added.
            this.ResumeLayout();
        }

        [STAThread]
        public static void Main(string[] args)
        {
            Application.Run(new Recipe07_01());
        }
    }
}
