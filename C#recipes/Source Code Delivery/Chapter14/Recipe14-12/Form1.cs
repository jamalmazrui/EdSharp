using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.WindowsAPICodePack.Dialogs;

namespace Recipe14_12
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        void showElevatedTaskRequest(object sender, EventArgs args)
        {
            // create the TaskDialog and configure the basics
            TaskDialog taskDialog = new TaskDialog();
            taskDialog.Cancelable = true;
            taskDialog.InstructionText = "This is a Task Dialog";
            taskDialog.StandardButtons = TaskDialogStandardButtons.Ok | TaskDialogStandardButtons.Close;

            // create the control that will represent the elevated task
            TaskDialogCommandLink commLink = new TaskDialogCommandLink(
                "adminTask", "First Line Of Text", "Second Line of Text");
            commLink.ShowElevationIcon = true;

            // add the control to the task dialog
            taskDialog.Controls.Add(commLink);

            // add an event handler to the command link so that
            // we are notified when the user presses the button
            commLink.Click += new EventHandler(performElevatedTask);

            // display the task dialog
            taskDialog.Show();
        }

        void performElevatedTask(object sender, EventArgs args)
        {
            textBox1.AppendText("Elevated task button pressed\n");
        }
    }
}
