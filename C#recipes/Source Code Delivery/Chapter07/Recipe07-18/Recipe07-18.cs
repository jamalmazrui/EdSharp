using System;
using System.Drawing;
using System.Windows.Forms;

namespace Apress.VisualCSharpRecipes.Chapter07
{
    public partial class Recipe07_18 : Form
    {
        public Recipe07_18()
        {
            // Initialization code is designer generated and contained
            // in a separate file named Recipe07-18.Designer.cs.
            InitializeComponent();

            this.richTextBox1.AllowDrop = true;
            this.richTextBox1.EnableAutoDragDrop = false;
            this.richTextBox1.DragDrop 
                += new System.Windows.Forms.DragEventHandler
                    (this.RichTextBox_DragDrop);
            this.richTextBox1.DragEnter 
                += new System.Windows.Forms.DragEventHandler
                    (this.RichTextBox_DragEnter);
        }

        private void RichTextBox_DragDrop(object sender, DragEventArgs e)
        {
            RichTextBox txt = sender as RichTextBox;

            if (txt != null)
            {
                // Insert the dragged text.
                int pos = txt.SelectionStart;
                
                string newText = txt.Text.Substring(0, pos)
                    + e.Data.GetData(DataFormats.Text).ToString()
                    + txt.Text.Substring(pos);

                txt.Text = newText;
            }
        }

        private void RichTextBox_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.Text))
            {
                e.Effect = DragDropEffects.Copy;
            }
            else
            {
                e.Effect = DragDropEffects.None;
            }
        }

        private void RichTextBox_MouseDown(object sender, MouseEventArgs e)
        {
            RichTextBox txt = sender as RichTextBox;

            // If the left mouse button is pressed and text is selected
            // this is a possible drag event.
            if (sender != null && txt.SelectionLength > 0
                && Form.MouseButtons == MouseButtons.Left)
            {
                // Only initiate a drag if the cursor is currently inside
                // the region of selected text.
                int pos = txt.GetCharIndexFromPosition(e.Location);

                if (pos >= txt.SelectionStart
                    && pos <= (txt.SelectionStart + txt.SelectionLength))
                {
                    txt.DoDragDrop(txt.SelectedText, DragDropEffects.Copy);
                }
            }
        }

        private void TextBox_DragDrop(object sender, DragEventArgs e)
        {
            TextBox txt = sender as TextBox;

            if (txt != null)
            {
                txt.Text = (string)e.Data.GetData(DataFormats.Text);
            }
        }

        private void TextBox_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.Text))
            {
                e.Effect = DragDropEffects.Copy;
            }
            else
            {
                e.Effect = DragDropEffects.None;
            }
        }

        private void TextBox_MouseDown(object sender, MouseEventArgs e)
        {
            TextBox txt = sender as TextBox;

            if (txt != null && Form.MouseButtons == MouseButtons.Left)
            {
                txt.SelectAll();
                txt.DoDragDrop(txt.Text, DragDropEffects.Copy);
            }
        }

        [STAThread]
        public static void Main(string[] args)
        {
            Application.Run(new Recipe07_18());
        }
    }
}