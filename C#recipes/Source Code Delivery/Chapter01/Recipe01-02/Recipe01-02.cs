using System;
using System.Windows.Forms;

namespace Apress.VisualCSharpRecipes.Chapter01
{
    class Recipe01_02 : Form
    {
        // Private members to hold references to the form's controls.
        private Label label1;
        private TextBox textBox1;
        private Button button1;

        // Constructor used to create an instance of the form and configure
        // the form's controls.
        public Recipe01_02() 
        {
            // Instantiate the controls used on the form.
            this.label1 = new Label();
            this.textBox1 = new TextBox();
            this.button1 = new Button();

            // Suspend the layout logic of the form while we configure and 
            // position the controls.
            this.SuspendLayout();

            // Configure label1, which displays the user prompt.
            this.label1.Location = new System.Drawing.Point(16, 36);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(148, 16);
            this.label1.TabIndex = 0;
            this.label1.Text = "Please enter your name:";

            // Configure textBox1, which accepts the user input.
            this.textBox1.Location = new System.Drawing.Point(172, 32);
            this.textBox1.Name = "textBox1";
            this.textBox1.TabIndex = 1;
            this.textBox1.Text = "";

            // Configure button1, which the user clicks to enter a name.
            this.button1.Location = new System.Drawing.Point(109, 80);
            this.button1.Name = "button1";
            this.button1.TabIndex = 2;
            this.button1.Text = "Enter";
            this.button1.Click += new System.EventHandler(this.button1_Click);

            // Configure WelcomeForm and add controls.
            this.ClientSize = new System.Drawing.Size(292, 126);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.textBox1);
            this.Controls.Add(this.label1);
            this.Name = "form1";
            this.Text = "Visual C# 2010 Recipes";

            // Resume the layout logic of the form now that all controls are 
            // configured.
            this.ResumeLayout(false);    
        }

        // Event handler called when the user clicks the Enter button on the 
        // form.
        private void button1_Click(object sender, System.EventArgs e)
        {
            // Write debug message to the console.
            System.Console.WriteLine("User entered: " + textBox1.Text);

            // Display welcome as a message box.
            MessageBox.Show("Welcome to Visual C# 2010 Recipes, "
                + textBox1.Text, "Visual C# 2010 Recipes");
        }

        // Application entry point, creates an instance of the form, and begins
        // running a standard message loop on the current thread. The message 
        // loop feeds the application with input from the user as events.
        [STAThread]
        public static void Main() 
        {
            Application.EnableVisualStyles();
            Application.Run(new Recipe01_02());
        }
    }
}
