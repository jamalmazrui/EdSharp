using System;
using System.IO;
using System.Windows.Forms;

namespace Apress.VisualCSharpRecipes.Chapter07
{
    public partial class Recipe07_02 : Form
    {
        public Recipe07_02()
        {
            // Initialization code is designer generated and contained
            // in a separate file named Recipe07-02.Designer.cs.
            InitializeComponent();
        }

        protected override void OnLoad(EventArgs e)
        {
            // Call the OnLoad method of the base class to ensure the Load 
            // event is raised correctly.
            base.OnLoad(e);

            // Get all the files in the root directory.
            DirectoryInfo directory = new DirectoryInfo(@"C:\");
            FileInfo[] files = directory.GetFiles();

            // Display the name of each file in the ListView.
            foreach (FileInfo file in files)
            {
                ListViewItem item = listView1.Items.Add(file.Name);
                item.ImageIndex = 0;

                // Associate each FileInfo object with its ListViewItem.
                item.Tag = file;
            }
        }

        private void listView1_ItemActivate(object sender, EventArgs e)
        {
            // Get information from the linked FileInfo object and display
            // it using MessageBox.
            ListViewItem item = ((ListView)sender).SelectedItems[0];
            FileInfo file = (FileInfo)item.Tag;
            string info = file.FullName + " is " + file.Length + " bytes.";

            MessageBox.Show(info, "File Information");
        }

        [STAThread]
        public static void Main(string[] args)
        {
            Application.Run(new Recipe07_02());
        }
    }
}