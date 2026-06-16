using System.Windows;
using System.Windows.Controls;

namespace Apress.VisualCSharpRecipes.Chapter17
{
    public class AlternatingRowStyleSelector : StyleSelector
    {
        // Flag to track the alternate rows
        private bool isAlternate = false;

        public Style DefaultStyle { get; set; }
        public Style AlternateStyle { get; set; }

        public override Style SelectStyle(object item, DependencyObject container)
        {
            // Select the style, based on the value of isAlternate
            Style style = isAlternate ? AlternateStyle : DefaultStyle;

            // Invert the flag
            isAlternate = !isAlternate;

            return style;
        }
    }
}
