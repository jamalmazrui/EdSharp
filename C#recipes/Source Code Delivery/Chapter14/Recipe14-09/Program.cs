using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

using Microsoft.WindowsAPICodePack.Taskbar;

namespace Recipe14_09
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {

            if (!TaskbarManager.IsPlatformSupported)
            {
                MessageBox.Show("This recipe requires Windows 7", "Recipe needs Windows 7", MessageBoxButtons.OK, MessageBoxIcon.Error);
                System.Environment.Exit(0);
                return;
            }

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Form1());
        }
    }
}
