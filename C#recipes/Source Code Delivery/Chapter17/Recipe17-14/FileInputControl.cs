using System.Windows;
using System.Windows.Controls;
using System.Windows.Markup;
using Microsoft.Win32;

namespace Apress.VisualCSharpRecipes.Chapter17
{
    [TemplatePart(Name = "PART_Browse", Type = typeof(Button))]
    [ContentProperty("FileName")]
    public class FileInputControl : Control
    {
        static FileInputControl()
        {
            DefaultStyleKeyProperty.OverrideMetadata(
                typeof(FileInputControl),
                new FrameworkPropertyMetadata(
                    typeof(FileInputControl)));
        }

        public override void OnApplyTemplate()
        {
            base.OnApplyTemplate();

            Button browseButton = base.GetTemplateChild("PART_Browse") as Button;

            if (browseButton != null)
                browseButton.Click += new RoutedEventHandler(browseButton_Click);
        }

        void browseButton_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog dlg = new OpenFileDialog();
            if (dlg.ShowDialog() == true)
            {
                this.FileName = dlg.FileName;
            }
        }

        public string FileName
        {
            get
            {
                return (string)GetValue(FileNameProperty);
            }
            set
            {
                SetValue(FileNameProperty, value);
            }
        }

        public static readonly DependencyProperty FileNameProperty =
            DependencyProperty.Register( "FileName", typeof(string), typeof(FileInputControl));
    }
}

