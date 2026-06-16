using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace Apress.VisualCSharpRecipes.Chapter12
{
    static class Recipe12_02
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new ActiveWindowInfo());
        }
    }
}