using System;
using System.Drawing;
using System.Windows.Forms;

namespace Apress.VisualCSharpRecipes.Chapter07
{
    public partial class Recipe07_15 : Form
    {
        // Boolean member tracks whether the form is in drag mode. If it is, 
        // mouse movements over the label will be translated into form movements.
        private bool dragging;

        // Stores the offset where the label is clicked.
        private Point pointClicked;

        public Recipe07_15()
        {
            // Initialization code is designer generated and contained
            // in a separate file named Recipe07-15.Designer.cs.
            InitializeComponent();
        }

        // MouseDown event handler for the label initiates the dragging process.
        private void lblDrag_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                // Turn drag mode on and store the point clicked.
                dragging = true;
                pointClicked = new Point(e.X, e.Y);
            }
            else
            {
                dragging = false;
            }
        }

        // MouseMove event handler for the label processes dragging movements if
        // the form is in drag mode.
        private void lblDrag_MouseMove(object sender, MouseEventArgs e)
        {
            if (dragging)
            {
                Point pointMoveTo;

                // Find the current mouse position in screen coordinates.
                pointMoveTo = this.PointToScreen(new Point(e.X, e.Y));

                // Compensate for the position the control was clicked.
                pointMoveTo.Offset(-pointClicked.X, -pointClicked.Y);

                // Move the form.
                this.Location = pointMoveTo;
            }   
        }

        // MouseUp event handler for the label switches of drag mode.
        private void lblDrag_MouseUp(object sender, MouseEventArgs e)
        {
            dragging = false;
        }

        private void cmdClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        [STAThread]
        public static void Main(string[] args)
        {
            Application.Run(new Recipe07_15());
        }
    }
}