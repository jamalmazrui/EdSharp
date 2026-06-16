using System;
using System.Drawing;
using System.Windows.Forms;

namespace Apress.VisualCSharpRecipes.Chapter07
{
    public partial class Recipe07_16 : Form
    {
        // An array to hold the set of Icons used to create the
        // animation effect.
        private Icon[] images = new Icon[8];

        // An integer to identify the current icon to display.
        private int offset = 0;

        public Recipe07_16()
        {
            // Initialization code is designer generated and contained
            // in a separate file named Recipe07-16.Designer.cs.
            InitializeComponent();

            // Declare a ContextMenuStrip for use by the NotifyIcon.
            ContextMenuStrip contextMenuStrip = new ContextMenuStrip();
            contextMenuStrip.Items.Add(new ToolStripMenuItem("About..."));
            contextMenuStrip.Items.Add(new ToolStripSeparator());
            contextMenuStrip.Items.Add(new ToolStripMenuItem("Exit"));

            // Assign the ContextMenuStrip to the NotifyIcon.
            notifyIcon.ContextMenuStrip = contextMenuStrip;
        }

        protected override void OnLoad(EventArgs e)
        {
            // Call the OnLoad method of the base class to ensure the Load 
            // event is raised correctly.
            base.OnLoad(e);

            // Load the basic set of eight icons.
            images[0] = new Icon("moon01.ico");
            images[1] = new Icon("moon02.ico");
            images[2] = new Icon("moon03.ico");
            images[3] = new Icon("moon04.ico");
            images[4] = new Icon("moon05.ico");
            images[5] = new Icon("moon06.ico");
            images[6] = new Icon("moon07.ico");
            images[7] = new Icon("moon08.ico");
        }

        private void timer_Elapsed(object sender, System.Timers.ElapsedEventArgs e)
        {
            // Change the icon. This event handler fires once every second 
            // (1000 ms).
            notifyIcon.Icon = images[offset];
            offset++;
            if (offset > 7) offset = 0;
        }

        [STAThread]
        public static void Main(string[] args)
        {
            Application.Run(new Recipe07_16());
        }
    }
}