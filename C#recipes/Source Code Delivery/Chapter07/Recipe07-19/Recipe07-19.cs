using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace Apress.VisualCSharpRecipes.Chapter07
{
    public partial class Recipe07_19 : Form
    {
        // Declare timers that change the button color.
        System.Timers.Timer greenTimer;
        System.Timers.Timer redTimer;

        public Recipe07_19()
        {
            // Initialization code is designer generated and contained
            // in a separate file named Recipe07-19.Designer.cs.
            InitializeComponent();

            // Create auto reset timers that fire at varying intervals 
            // to change the color of the button on the form.
            greenTimer = new System.Timers.Timer(3000);
            greenTimer.Elapsed += 
                new System.Timers.ElapsedEventHandler(greenTimer_Elapsed);
            greenTimer.Start();

            redTimer = new System.Timers.Timer(5000);
            redTimer.Elapsed += 
                new System.Timers.ElapsedEventHandler(redTimer_Elapsed);
            redTimer.Start();
        }

        void redTimer_Elapsed(object sender, System.Timers.ElapsedEventArgs e)
        {
            // Use an anonymous method to set the button color to red.
            button1.Invoke((Action)delegate {button1.BackColor = Color.Red;});
        }

        void greenTimer_Elapsed(object sender, System.Timers.ElapsedEventArgs e)
        {
            // Use a lambda expression to set the button color to green.
            button1.Invoke(new Action(() => button1.BackColor = Color.Green));
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        [STAThread]
        public static void Main(string[] args)
        {
            Application.Run(new Recipe07_19());
        }
    }
}