using System;
using System.Drawing;
using System.Windows.Forms;

namespace Apress.VisualCSharpRecipes.Chapter08
{
    public partial class Recipe08_04 : Form
    {
        public Recipe08_04()
        {
            InitializeComponent();
        }

        // Tracks when drag mode is on.
        private bool isDraggingA = false;
        private bool isDraggingB = false;

        // The ellipse shape controls.
        private EllipseShape ellipseA, ellipseB;

        private void Recipe08_04_Load(object sender, EventArgs e)
        {
            // Create and configure both ellipses.
            ellipseA = new EllipseShape();
            ellipseA.Width = ellipseA.Height = 100;
            ellipseA.Top = ellipseA.Left = 30;
            ellipseA.BackColor = Color.Red;
            this.Controls.Add(ellipseA);

            ellipseB = new EllipseShape();
            ellipseB.Width = ellipseB.Height = 100;
            ellipseB.Top = ellipseB.Left = 130;
            ellipseB.BackColor = Color.Azure;
            this.Controls.Add(ellipseB);

            // Attach both ellipses to the same set of event handlers.
            ellipseA.MouseDown += new MouseEventHandler(Ellipse_MouseDown);
            ellipseA.MouseUp += new MouseEventHandler(Ellipse_MouseUp);
            ellipseA.MouseMove += new MouseEventHandler(Ellipse_MouseMove);

            ellipseB.MouseDown += new MouseEventHandler(Ellipse_MouseDown);
            ellipseB.MouseUp += new MouseEventHandler(Ellipse_MouseUp);
            ellipseB.MouseMove += new MouseEventHandler(Ellipse_MouseMove);
        }

        private void Ellipse_MouseDown(object sender, MouseEventArgs e)
        {
            // Get the ellipse that triggered this event.
            Control control = (Control)sender;

            if (e.Button == MouseButtons.Left)
            {
                control.Tag = new Point(e.X, e.Y);
                if (control == ellipseA)
                {
                    isDraggingA = true;
                }
                else
                {
                    isDraggingB = true;
                }
            }
        }

        private void Ellipse_MouseUp(object sender, MouseEventArgs e)
        {
            isDraggingA = false;
            isDraggingB = false;
        }

        private void Ellipse_MouseMove(object sender, MouseEventArgs e)
        {
            // Get the ellipse that triggered this event.
            Control control = (Control)sender;

            if ((isDraggingA && control == ellipseA) ||
             (isDraggingB && control == ellipseB))
            {
                // Get the offset.
                Point point = (Point)control.Tag;

                // Move the control.
                control.Left = e.X + control.Left - point.X;
                control.Top = e.Y + control.Top - point.Y;
            }
        }
    }
}