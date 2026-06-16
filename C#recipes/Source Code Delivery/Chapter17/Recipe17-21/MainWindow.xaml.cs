using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace Apress.VisualCSharpRecipes.Chapter17
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private ListBoxItem draggedItem;
        private Point startDragPoint;

        public MainWindow()
        {
            InitializeComponent();
        }

        // Handles the DragEnter event for the Canvas. Changes the mouse
        // pointer to show the user that copy is an option if the drop
        // text content is over the Canvas.
        private void cvsSurface_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.Text))
            {
                e.Effects = DragDropEffects.Copy;
            }
            else
            {
                e.Effects = DragDropEffects.None;
            }
        }

        // Handles the Drop event for the Canvas. Creates a new Label
        // and adds it to the Canvas at the location of the mouse pointer.
        private void cvsSurface_Drop(object sender, DragEventArgs e)
        {
            // Create a new Label.
            Label newLabel = new Label();
            newLabel.Content = e.Data.GetData(DataFormats.Text);
            newLabel.FontSize = 14;

            // Add the Label to the Canvas and position it.
            cvsSurface.Children.Add(newLabel);
            Canvas.SetLeft(newLabel, e.GetPosition(cvsSurface).X);
            Canvas.SetTop(newLabel, e.GetPosition(cvsSurface).Y);
        }

        // Handles the PreviewMouseLeftButtonDown event for all ListBoxItem
        // objects. Stores a reference to the item being dragged and the
        // point at which the drag started.
        private void ListBoxItem_PreviewMouseLeftButtonDown(object sender,
            MouseButtonEventArgs e)
        {
            draggedItem = sender as ListBoxItem;
            startDragPoint = e.GetPosition(null);
        }

        // Handles the PreviewMouseMove event for all ListBoxItem objects.
        // Determines whether the mouse has been moved far enough to be
        // considered a drag operation.
        private void ListBoxItem_PreviewMouseMove(object sender,
            MouseEventArgs e)
        {
            if (e.LeftButton == MouseButtonState.Pressed)
            {
                Point position = e.GetPosition(null);

                if (Math.Abs(position.X - startDragPoint.X) >
                        SystemParameters.MinimumHorizontalDragDistance ||
                    Math.Abs(position.Y - startDragPoint.Y) >
                        SystemParameters.MinimumVerticalDragDistance)
                {
                    // User is dragging, set up the DragDrop behavior.
                    DragDrop.DoDragDrop(draggedItem, draggedItem.Content,
                        DragDropEffects.Copy);
                }
            }
        }
    }
}
