using System;
using System.Windows.Forms;

namespace Apress.VisualCSharpRecipes.Chapter05
{
    static class Recipe05_06
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new DirectoryTree());
        }
    }
}