using System;
using System.IO;
using System.Drawing;
using System.Windows.Forms;

namespace Apress.VisualCSharpRecipes.Chapter07
{
    public partial class Recipe07_09 : Form
    {
        public Recipe07_09()
        {
            // Initialization code is designer generated and contained
            // in a separate file named Recipe07-09.Designer.cs.
            InitializeComponent();

            // Configure ComboBox1 to make its autocomplete
            // suggestions from a custom source.
            this.comboBox1.AutoCompleteCustomSource.AddRange(
                new string[] { "Man", "Mark", "Money", "Motley", 
                    "Mostly", "Mint", "Minion", "Milk", "Mist", 
                    "Mush", "More", "Map", "Moon", "Monkey"});

            this.comboBox1.AutoCompleteMode 
                = AutoCompleteMode.SuggestAppend;
            this.comboBox1.AutoCompleteSource 
                = AutoCompleteSource.CustomSource;

            // Configure ComboBox2 to make its autocomplete
            // suggestions from its current contents.
            this.comboBox2.Items.AddRange(
                new string[] { "Man", "Mark", "Money", "Motley", 
                    "Mostly", "Mint", "Minion", "Milk", "Mist", 
                    "Mush", "More", "Map", "Moon", "Monkey"});

            this.comboBox2.AutoCompleteMode
                = AutoCompleteMode.SuggestAppend;
            this.comboBox2.AutoCompleteSource
                = AutoCompleteSource.ListItems;

            // Configure ComboBox3 to make its autocomplete
            // suggestions from the systems URL history.
            this.comboBox3.AutoCompleteMode
                = AutoCompleteMode.SuggestAppend;
            this.comboBox3.AutoCompleteSource
                = AutoCompleteSource.AllUrl;
        }

        [STAThread]
        public static void Main(string[] args)
        {
            Application.Run(new Recipe07_09());
        }
    }
}