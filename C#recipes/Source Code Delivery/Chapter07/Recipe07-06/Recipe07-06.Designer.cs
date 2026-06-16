namespace Apress.VisualCSharpRecipes.Chapter07
{
    partial class Recipe07_06
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.redButton = new System.Windows.Forms.Button();
            this.blueButton = new System.Windows.Forms.Button();
            this.greenButton = new System.Windows.Forms.Button();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.SuspendLayout();
            // 
            // redButton
            // 
            this.redButton.ForeColor = System.Drawing.Color.Red;
            this.redButton.Location = new System.Drawing.Point(12, 231);
            this.redButton.Name = "redButton";
            this.redButton.Size = new System.Drawing.Size(75, 23);
            this.redButton.TabIndex = 0;
            this.redButton.Text = "Red";
            this.redButton.Click += new System.EventHandler(this.Button_Click);
            // 
            // blueButton
            // 
            this.blueButton.ForeColor = System.Drawing.Color.Blue;
            this.blueButton.Location = new System.Drawing.Point(108, 231);
            this.blueButton.Name = "blueButton";
            this.blueButton.Size = new System.Drawing.Size(75, 23);
            this.blueButton.TabIndex = 2;
            this.blueButton.Text = "Blue";
            this.blueButton.Click += new System.EventHandler(this.Button_Click);
            // 
            // greenButton
            // 
            this.greenButton.ForeColor = System.Drawing.Color.Green;
            this.greenButton.Location = new System.Drawing.Point(205, 231);
            this.greenButton.Name = "greenButton";
            this.greenButton.Size = new System.Drawing.Size(75, 23);
            this.greenButton.TabIndex = 3;
            this.greenButton.Text = "Green";
            this.greenButton.Click += new System.EventHandler(this.Button_Click);
            // 
            // textBox1
            // 
            this.textBox1.BackColor = global::Apress.VisualCSharpRecipes.Chapter07.Properties.Settings.Default.Color;
            this.textBox1.DataBindings.Add(new System.Windows.Forms.Binding("BackColor", global::Apress.VisualCSharpRecipes.Chapter07.Properties.Settings.Default, "Color", true, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged));
            this.textBox1.Location = new System.Drawing.Point(12, 12);
            this.textBox1.Multiline = true;
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(268, 213);
            this.textBox1.TabIndex = 1;
            this.textBox1.Text = "This recipe demonstrates the use of .NET 2.0 Application Settings.";
            // 
            // Recipe07_06
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(292, 266);
            this.Controls.Add(this.greenButton);
            this.Controls.Add(this.blueButton);
            this.Controls.Add(this.textBox1);
            this.Controls.Add(this.redButton);
            this.DataBindings.Add(new System.Windows.Forms.Binding("Size", global::Apress.VisualCSharpRecipes.Chapter07.Properties.Settings.Default, "Size", true, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged));
            this.DataBindings.Add(new System.Windows.Forms.Binding("Location", global::Apress.VisualCSharpRecipes.Chapter07.Properties.Settings.Default, "Location", true, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged));
            this.Location = global::Apress.VisualCSharpRecipes.Chapter07.Properties.Settings.Default.Location;
            this.Name = "Recipe07_06";
            this.Text = "Recipe07_06";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button redButton;
        private System.Windows.Forms.Button blueButton;
        private System.Windows.Forms.Button greenButton;
        private System.Windows.Forms.TextBox textBox1;
    }
}