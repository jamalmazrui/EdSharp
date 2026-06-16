using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.IO;
using System.Windows.Forms;

using Microsoft.WindowsAPICodePack.Shell;
using Microsoft.WindowsAPICodePack.Taskbar;

namespace Recipe14_09
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            // call the default intializer
            InitializeComponent();

            // register an event handler for when the windows is shown
            this.Shown += new EventHandler(onWindowShown);
        }

        void onWindowShown(object sender, EventArgs e)
        {
            // create a new jump list
            JumpList jumpList = JumpList.CreateJumpList();

            // create a user task
            jumpList.AddUserTasks(new JumpListLink("cmd.exe", "Open Command Prompt"));

            // create a user task with an icon
            jumpList.AddUserTasks(new JumpListLink("notepad.exe", "Open Notepad")
            {
                IconReference = new IconReference("notepad.exe", 0)
            });

            // create a user task with a document
            jumpList.AddUserTasks(new JumpListLink(@"C:\Users\Adam\Desktop\test.txt", 
                "Open Text File"));
            
        }
    }
}
