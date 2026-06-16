using System;
using System.ComponentModel;
using System.Windows.Forms;

namespace Apress.VisualCSharpRecipes.Chapter07
{
    public partial class Recipe07_21 : Form
    {
        private Window1 modelessWindow;
        private CancelEventHandler modelessWindowCloseHandler;

        public Recipe07_21()
        {
            // Initialization code is designer generated and contained
            // in a separate file named Recipe07-21.Designer.cs.
            InitializeComponent();
            modelessWindowCloseHandler = new CancelEventHandler(Window_Closing);
        }

        // Handles the button click event to open and close the modeless 
        // WPF Window.
        private void OpenModeless_Click(object sender, EventArgs e)
        {
            if (modelessWindow == null)
            {
                modelessWindow = new Window1();

                // Add an event handler to get notification when the window
                // is closing.
                modelessWindow.Closing += modelessWindowCloseHandler;

                // Change the button text.
                btnOpenModeless.Text = "Close Modeless Window";

                // Show the Windows Form.
                modelessWindow.Show();
            }
            else
            {
                modelessWindow.Close();
            }
        }

        // Handles the button click event to open the modal WPF Window.
        private void OpenModal_Click(object sender, EventArgs e)
        {
            // Create and display the modal window.
            Window1 window = new Window1();
            window.ShowDialog();
        }

        // Handles the WPF Window's Closing event for the modeless window.
        private void Window_Closing(object sender, CancelEventArgs e)
        {
            // Remove the event handler reference.
            modelessWindow.Closing -= modelessWindowCloseHandler;
            modelessWindow = null;

            // Change the button text.
            btnOpenModeless.Text = "Open Modeless Window";
        }
    }
}
