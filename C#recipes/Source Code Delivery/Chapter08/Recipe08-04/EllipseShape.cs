using System;
using System.Drawing;
using System.Windows.Forms;
using System.Drawing.Drawing2D;

namespace Apress.VisualCSharpRecipes.Chapter08
{
    public partial class EllipseShape : Control
    {
        public EllipseShape()
        {
            InitializeComponent();
        }

        private GraphicsPath path = null;

        private void RefreshPath()
        {
            // Create the GraphicsPath for the shape (in this case
            // an ellipse that fits inside the full control area)
            // and apply it to the control by setting
            // the Region property.
            path = new GraphicsPath();
            path.AddEllipse(this.ClientRectangle);
            this.Region = new Region(path);
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            base.OnPaint(e);
            if (path != null)
            {
                e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
                e.Graphics.FillPath(new SolidBrush(this.BackColor), path);
                e.Graphics.DrawPath(new Pen(this.ForeColor, 4), path);
            }
        }

        protected override void OnResize(System.EventArgs e)
        {
            base.OnResize(e);
            RefreshPath();
            this.Invalidate();
        }
    }
}
