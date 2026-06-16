using System;
using System.Windows.Forms;
using System.Media;
using System.Threading;

namespace Apress.VisualCSharpRecipes.Chapter08
{
    public partial class Recipe08_09 : Form
    {
        public Recipe08_09()
        {
            InitializeComponent();
        }

        private void Recipe08_09_Load(object sender, EventArgs e)
        {
            // Play a beep with default frequency 
            // and duration (800 and 200 respectively)
            Console.Beep();
            Thread.Sleep(500);

            // Play a beep with frequency as 200 and duration as 300
            Console.Beep(200, 300);
            Thread.Sleep(500);

            // Play the sound associated with the Asterisk event
            SystemSounds.Asterisk.Play();
        }
    }
}