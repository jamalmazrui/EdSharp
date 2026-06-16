using System;
using System.Windows;
using System.Windows.Forms;
using System.Windows.Forms.Integration;
using WPFControls = System.Windows.Controls;

namespace Apress.VisualCSharpRecipes.Chapter07
{
    public partial class Recipe07_22 : Form
    {
        WPFControls.TextBox textBox;

        public Recipe07_22()
        {
            // Initialization code is designer generated and contained
            // in a separate file named Recipe07-21.Designer.cs.
            InitializeComponent();

            // Create a new WPF TextBox control.
            textBox = new WPFControls.TextBox();
            textBox.Text = "A WPF TextBox\n\r\n\r";
            textBox.TextAlignment = TextAlignment.Center;
            textBox.VerticalAlignment = VerticalAlignment.Center;
            textBox.VerticalScrollBarVisibility =
                WPFControls.ScrollBarVisibility.Auto;
            textBox.IsReadOnly = true;

            // Create a new ElementHost to host the WPF TextBox.
            ElementHost elementHost2 = new ElementHost();
            elementHost2.Name = "elementHost2";
            elementHost2.Dock = DockStyle.Fill;
            elementHost2.Child = textBox;
            elementHost2.Size = new System.Drawing.Size(156, 253);
            elementHost2.RightToLeft = RightToLeft.No;

            // Place the new ElementHost in the bottom left table cell.
            tableLayoutPanel1.Controls.Add(elementHost2, 1, 0);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Change the ellipse color.
            ellipseControl1.ChangeColor();

            // Get the current ellipse color and append to TextBox.
            textBox.Text +=
                String.Format("Ellipse color changed to {0}\n\r",
                ellipseControl1.Color);

            textBox.ScrollToEnd();
        }
    }
}