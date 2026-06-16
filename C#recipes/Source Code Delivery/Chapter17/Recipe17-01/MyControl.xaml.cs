using System.Windows;
using System.Windows.Controls;

namespace Apress.VisualCSharpRecipes.Chapter17
{
    public partial class MyControl : UserControl
    {
        public MyControl()
        {
            InitializeComponent();
            DataContext = this;
        }

        public FontWeight TextFontWeight
        {
            get { return (FontWeight)GetValue(TextFontWeightProperty); }
            set { SetValue(TextFontWeightProperty, value); }
        }

        public static readonly DependencyProperty TextFontWeightProperty =
            DependencyProperty.Register("TextFontWeight", typeof(FontWeight), typeof(MyControl), 
                                        new FrameworkPropertyMetadata(FontWeights.Normal, 
                                                                      FrameworkPropertyMetadataOptions.AffectsArrange 
                                                                      & FrameworkPropertyMetadataOptions.AffectsMeasure 
                                                                      & FrameworkPropertyMetadataOptions.AffectsRender, 
                                                                      TextFontWeight_PropertyChanged, 
                                                                      TextFontWeight_CoerceValue));

        public string TextContent
        {
            get { return (string)GetValue(TextContentProperty); }
            set { SetValue(TextContentProperty, value); }
        }

        public static readonly DependencyProperty TextContentProperty =
            DependencyProperty.Register("TextContent", typeof(string), typeof(MyControl), 
                                        new FrameworkPropertyMetadata("Default Value", 
                                                                      FrameworkPropertyMetadataOptions.AffectsArrange 
                                                                      & FrameworkPropertyMetadataOptions.AffectsMeasure
                                                                      & FrameworkPropertyMetadataOptions.AffectsRender));

        private static object TextFontWeight_CoerceValue(DependencyObject d, object value)
        {
            FontWeight fontWeight = (FontWeight)value;

            if (fontWeight == FontWeights.Bold
                || fontWeight == FontWeights.Normal)
            {
                return fontWeight;
            }

            return FontWeights.Normal;
        }

        private static void TextFontWeight_PropertyChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            MyControl myControl = d as MyControl;

            if (myControl != null)
            {
                FontWeight fontWeight = (FontWeight)e.NewValue;
                string fontWeightName;

                if (fontWeight == FontWeights.Bold)
                    fontWeightName = "Bold";
                else
                    fontWeightName = "Normal";

                myControl.txblFontWeight.Text = string.Format("Font weight set to: {0}.", fontWeightName);
            }
        }
    }
}
