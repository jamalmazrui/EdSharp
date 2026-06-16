using System;
using System.Drawing;
using System.Windows.Forms;

namespace Apress.VisualCSharpRecipes.Chapter07
{
    public partial class Recipe07_12 : Form
    {
        public Recipe07_12()
        {
            // Initialization code is designer generated and contained
            // in a separate file named Recipe07-12.Designer.cs.
            InitializeComponent();
        }

        // As the main form loads, clone the required section of the main 
        // menu and assign it to the ContextMenu property of the TextBox.
        protected override void OnLoad(EventArgs e)
        {
            // Call the OnLoad method of the base class to ensure the Load 
            // event is raised correctly.
            base.OnLoad(e);

            ContextMenu mnuContext = new ContextMenu();

            // Copy the menu items from the File menu into a context menu.
            foreach (MenuItem mnuItem in mnuFile.MenuItems)
            {
                mnuContext.MenuItems.Add(mnuItem.CloneMenu());
            }

            // Attach the cloned menu to the TextBox.
            TextBox1.ContextMenu = mnuContext;
        }

        // Event handler to display the ContextMenu for the ListBox.
        private void TextBox1_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                TextBox1.ContextMenu.Show(TextBox1, new Point(e.X, e.Y));
            }
        }

        // Event handler to process clicks on File/Open menu item.
        // For the purpose of the example, simply show a message box.
        private void mnuOpen_Click(object sender, EventArgs e)
        {
            MessageBox.Show("This is the event handler for Open.","Recipe07-12");
        }

        // Event handler to process clicks on File/Save menu item.
        // For the purpose of the example, simply show a message box.
        private void mnuSave_Click(object sender, EventArgs e)
        {
            MessageBox.Show("This is the event handler for Save.","Recipe07-12");
        }

        // Event handler to process clicks on File/Exit menu item.
        // For the purpose of the example, simply show a message box.
        private void mnuExit_Click(object sender, EventArgs e)
        {
            MessageBox.Show("This is the event handler for Exit.","Recipe07-12");
        }

        [STAThread]
        public static void Main(string[] args)
        {
            Application.Run(new Recipe07_12());
        }
    }
}