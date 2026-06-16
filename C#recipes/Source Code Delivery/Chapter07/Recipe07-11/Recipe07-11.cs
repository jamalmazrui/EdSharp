using System;
using System.Drawing;
using System.Windows.Forms;

namespace Apress.VisualCSharpRecipes.Chapter07
{
    public partial class Recipe07_11 : Form
    {
        public Recipe07_11()
        {
            // Initialization code is designer generated and contained
            // in a separate file named Recipe07-11.Designer.cs.
            InitializeComponent();
        }

        [STAThread]
        public static void Main(string[] args)
        {
            Application.Run(new Recipe07_11());
        }
    }
}