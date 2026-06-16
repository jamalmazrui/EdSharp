using System;
using System.Drawing;
using System.Windows.Forms;
using System.Drawing.Drawing2D;

namespace Apress.VisualCSharpRecipes.Chapter08
{
    public partial class Recipe08_02 : Form
    {
        // Define the shapes used on this form.
        private GraphicsPath path;
        private Rectangle rectangle;

        // Define the flags that track where the mouse pointer is.
        private bool inPath = false;
        private bool inRectangle = false;

        // Define the brushes used for painting the shapes.
        Brush highlightBrush = Brushes.HotPink;
        Brush defaultBrush = Brushes.LightBlue;

        public Recipe08_02()
        {
            InitializeComponent();
        }

        private void Recipe08_02_Load(object sender, EventArgs e)
        {
            // Create the shapes that will be displayed.
            path = new GraphicsPath();
            path.AddEllipse(10, 10, 100, 60);
            path.AddCurve(new Point[] {new Point(50, 50),
                      new Point(10,33), new Point(80,43)});
            path.AddLine(50, 120, 250, 80);
            path.AddLine(120, 40, 110, 50);
            path.CloseFigure();

            rectangle = new Rectangle(100, 170, 220, 120);
        }

        private void Recipe08_02_Paint(object sender, PaintEventArgs e)
        {
            Graphics g = e.Graphics;

            // Paint the shapes according to the current selection.
            if (inPath)
            {
                g.FillPath(highlightBrush, path);
                g.FillRectangle(defaultBrush, rectangle);
            }
            else if (inRectangle)
            {
                g.FillRectangle(highlightBrush, rectangle);
                g.FillPath(defaultBrush, path);
            }
            else
            {
                g.FillPath(defaultBrush, path);
                g.FillRectangle(defaultBrush, rectangle);
            }
            g.DrawPath(Pens.Black, path);
            g.DrawRectangle(Pens.Black, rectangle);
        }

        private void Recipe08_02_MouseMove(object sender, MouseEventArgs e)
        {
            using (Graphics g = this.CreateGraphics())
            {
                // Perform hit testing with rectangle.
                if (rectangle.Contains(e.X, e.Y))
                {
                    if (!inRectangle)
                    {
                        inRectangle = true;

                        // Highlight the rectangle.
                        g.FillRectangle(highlightBrush, rectangle);
                        g.DrawRectangle(Pens.Black, rectangle);
                    }
                }
                else if (inRectangle)
                {
                    inRectangle = false;

                    // Restore the unhighlighted rectangle.
                    g.FillRectangle(defaultBrush, rectangle);
                    g.DrawRectangle(Pens.Black, rectangle);
                }

                // Perform hit testing with path.
                if (path.IsVisible(e.X, e.Y))
                {
                    if (!inPath)
                    {
                        inPath = true;

                        // Highlight the path.
                        g.FillPath(highlightBrush, path);
                        g.DrawPath(Pens.Black, path);
                    }
                }
                else if (inPath)
                {
                    inPath = false;

                    // Restore the unhighlighted path.
                    g.FillPath(defaultBrush, path);
                    g.DrawPath(Pens.Black, path);
                }
            }
        }
    }
}