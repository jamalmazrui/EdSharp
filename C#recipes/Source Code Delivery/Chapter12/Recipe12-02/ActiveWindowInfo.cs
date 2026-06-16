using System;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.Text;

namespace Apress.VisualCSharpRecipes.Chapter12
{
    public partial class ActiveWindowInfo : Form
    { 
        public ActiveWindowInfo()
        {
            InitializeComponent();
        }

        // Declare external functions.
        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll")]
        private static extern int GetWindowText(IntPtr hWnd, 
            StringBuilder text, int count);

        private void tmrRefresh_Tick(object sender, EventArgs e)
        {
            int chars = 256;
            StringBuilder buff = new StringBuilder(chars);

            // Obtain the handle of the active window.
            IntPtr handle = GetForegroundWindow();

            // Update the controls.
            if (GetWindowText(handle, buff, chars) > 0)
            {
                lblCaption.Text = buff.ToString();
                lblHandle.Text = handle.ToString();
                if (handle == this.Handle)
                {
                    lblCurrent.Text = "True";
                }
                else
                {
                    lblCurrent.Text = "False";
                }
            }
        }
    }
}