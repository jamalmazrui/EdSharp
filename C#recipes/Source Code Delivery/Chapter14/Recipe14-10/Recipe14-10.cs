using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.WindowsAPICodePack.Shell;
using Microsoft.WindowsAPICodePack.Shell.PropertySystem;

namespace Recipe14_10
{
    public partial class Recipe14_10 : Form
    {
        public Recipe14_10()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // create the leaf condition for the file extension
            SearchCondition fileExtCondition = 
                SearchConditionFactory.CreateLeafCondition(
                SystemProperties.System.FileExtension, textBox1.Text, 
                SearchConditionOperation.Equal);

            // create the leaf condition for the file name
            SearchCondition fileNameCondition =
                SearchConditionFactory.CreateLeafCondition(
                SystemProperties.System.FileName, textBox2.Text, 
                SearchConditionOperation.ValueContains);

            // combine the two leaf conditions
            SearchCondition comboCondition =
                SearchConditionFactory.CreateAndOrCondition(
                SearchConditionType.And, 
                false, fileExtCondition, fileNameCondition);
            
            // create the search folder
            ShellSearchFolder searchFolder = new ShellSearchFolder(
                comboCondition, (ShellContainer)KnownFolders.UsersFiles);

            // clear the result text box
            textBox3.Clear();
            textBox3.AppendText("Processing search results...\n");

            // run through each search result 
            foreach (ShellObject shellObject in searchFolder)
            {
                textBox3.AppendText("Result: " 
                    + shellObject.ParsingName + "\n");
            }

            // display a final message to the user
            textBox3.AppendText("All results processed\n");

        }
    }
}
