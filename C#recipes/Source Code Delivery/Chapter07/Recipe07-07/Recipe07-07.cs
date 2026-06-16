using System;
using System.Windows.Forms;

namespace Apress.VisualCSharpRecipes.Chapter07
{
    public partial class Recipe07_07 : Form
    {
        // A counter to keep track of the number of items added
        // to the ListBox.
        private int counter = 0;
        
        public Recipe07_07()
        {
            // Initialization code is designer generated and contained
            // in a separate file named Recipe07-07.Designer.cs.
            InitializeComponent();
        }

        // Button click event handler adds 20 new items to the ListBox.
        private void cmdTest_Click(object sender, EventArgs e)
        {
            // Add 20 items.
            for (int i = 0; i < 20; i++)
            {
                counter++;
                listBox1.Items.Add("Item " + counter.ToString());
            }

            // Set the TopIndex property of the ListBox to ensure the 
            // most recently added items are visible.
            listBox1.TopIndex = listBox1.Items.Count - 1;
            listBox1.SelectedIndex = listBox1.Items.Count - 1;
        }

        [STAThread]
        public static void Main(string[] args)
        {
            Application.Run(new Recipe07_07());
        }
    }
}