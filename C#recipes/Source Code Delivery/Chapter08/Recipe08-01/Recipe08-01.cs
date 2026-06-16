using System;
using System.Drawing;
using System.Windows.Forms;
using System.Drawing.Text;

namespace Apress.VisualCSharpRecipes.Chapter08
{
    public partial class Recipe08_01 : Form
    {
        public Recipe08_01()
        {
            InitializeComponent();
        }

        private void Recipe08_01_Load(object sender, EventArgs e)
        {
            // Create the font collection.
            using (InstalledFontCollection fontFamilies = 
                new InstalledFontCollection())
            {
                // Iterate through all font families.
                int offset = 10;
                foreach (FontFamily family in fontFamilies.Families)
                {
                    try
                    {
                        // Create a label that will display text in this font.
                        Label fontLabel = new Label();
                        fontLabel.Text = family.Name;
                        fontLabel.Font = new Font(family, 14);
                        fontLabel.Left = 10;
                        fontLabel.Width = pnlFonts.Width;
                        fontLabel.Top = offset;

                        // Add the label to a scrollable Panel.
                        pnlFonts.Controls.Add(fontLabel);
                        offset += 30;
                    }
                    catch
                    {
                        // An error will occur if the selected font does
                        // not support normal style (the default used when
                        // creating a Font object). This problem can be
                        // harmlessly ignored.
                    }
                }
            }
        }
    }
}