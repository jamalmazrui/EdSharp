using System;
using System.Drawing;
using System.Windows.Forms;
using System.Drawing.Drawing2D;

namespace Apress.VisualCSharpRecipes.Chapter08
{
    public partial class Recipe08_03 : Form
    {
        public Recipe08_03()
        {
            InitializeComponent();
        }

        private void Recipe08_03_Load(object sender, EventArgs e)
        {
            GraphicsPath path = new GraphicsPath();

            Point[] pointsA = new Point[]
                {
                    new Point(0, 0),
                    new Point(40, 60), 
                    new Point(this.Width - 100, 10)
                };
            path.AddCurve(pointsA);

            Point[] pointsB = new Point[]
                {
                    new Point(this.Width - 40, this.Height - 60), 
                    new Point(this.Width, this.Height),
                    new Point(10, this.Height)
                };
            path.AddCurve(pointsB);

            path.CloseAllFigures();

            this.Region = new Region(path);
        }

        private void cmdClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}