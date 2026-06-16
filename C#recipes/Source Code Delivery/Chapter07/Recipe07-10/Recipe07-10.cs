using System;
using System.Collections;
using System.Windows.Forms;

namespace Apress.VisualCSharpRecipes.Chapter07
{
    public partial class Recipe07_10 : Form
    {
        public Recipe07_10()
        {
            // Initialization code is designer generated and contained
            // in a separate file named Recipe07-10.Designer.cs.
            InitializeComponent();
        }

        // Event handler to handle user clicks on column headings.
        private void listView1_ColumnClick(object sender, ColumnClickEventArgs e)
        {
            // Create and/or configure the ListViewItemComparer to sort based on
            // the column that was clicked. 
            ListViewItemComparer sorter = 
                listView1.ListViewItemSorter as ListViewItemComparer;

            if (sorter == null)
            {
                // Create a new ListViewItemComparer.
                sorter = new ListViewItemComparer(e.Column);
                listView1.ListViewItemSorter = sorter;
            }
            else
            {
                // Configure the existing ListViewItemComparer.
                sorter.Column = e.Column;
            }

            // Sort the ListView
            listView1.Sort();
        }

        [STAThread]
        public static void Main(string[] args)
        {
            Application.Run(new Recipe07_10());
        }
    }

    public class ListViewItemComparer : IComparer
    {
        // Property to get/set the column to use for comparison.
        public int Column { get; set; }

        // Property to get/set whether numeric comparison is required
        // as opposed to the standard alphabetic comparison.
        public bool Numeric { get; set; }

        public ListViewItemComparer(int columnIndex)
        {
            Column = columnIndex;
        }

        public int Compare(object x, object y)
        {
            // Convert the arguments to ListViewItem objects.
            ListViewItem itemX = x as ListViewItem;
            ListViewItem itemY = y as ListViewItem;

            // Handle logic for null reference as dictated by the 
            // IComparer interface. Null is considered less than
            // any other value.
            if (itemX == null && itemY == null) return 0;
            else if (itemX == null) return -1;
            else if (itemY == null) return 1;

            // Shortcircuit condition where the items are references
            // to the same object.
            if (itemX == itemY) return 0;

            // Determine if numeric comparison is required.
            if (Numeric)
            {
                // Convert column text to numbers before comparing.
                // If the conversion fails, just use the value 0.
                decimal itemXVal, itemYVal;

                if (!Decimal.TryParse(itemX.SubItems[Column].Text, out itemXVal))
                {
                    itemXVal = 0;
                }
                if (!Decimal.TryParse(itemY.SubItems[Column].Text, out itemYVal))
                {
                    itemYVal = 0;
                }

                return Decimal.Compare(itemXVal, itemYVal);
            }
            else
            {
                // Keep the column text in its native string format
                // and perform an alphabetic comparison.
                string itemXText = itemX.SubItems[Column].Text;
                string itemYText = itemY.SubItems[Column].Text;

                return String.Compare(itemXText, itemYText);
            }
        }
    }
}