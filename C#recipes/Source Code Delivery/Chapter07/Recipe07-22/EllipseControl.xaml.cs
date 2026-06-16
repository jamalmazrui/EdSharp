using System.Windows.Controls;
using System.Windows.Media;

namespace Apress.VisualCSharpRecipes.Chapter07
{
    /// <summary>
    /// Interaction logic for EllipseControl.xaml
    /// </summary>
    public partial class EllipseControl : UserControl
    {
        public EllipseControl()
        {
            InitializeComponent();
        }

        // Gets the name of the current color.
        public string Color
        {
            get
            {
                if (Ellipse1.Fill == (Brush)Grid1.Resources["RedBrush"])
                {
                    return "Red";
                }
                else
                {
                    return "Blue";
                }
            }
        }

        // Switch the fill to the red gradient.
        public void ChangeColor()
        {
            // Check the current fill of the ellipse.
            if (Ellipse1.Fill == (Brush)Grid1.Resources["RedBrush"])
            {
                // Ellipse is red, change to blue.
                Ellipse1.Fill = (Brush)Grid1.Resources["BlueBrush"];
            }
            else
            {
                // Ellipse is blue, change to red.
                Ellipse1.Fill = (Brush)Grid1.Resources["RedBrush"];
            }
        }
    }
}
