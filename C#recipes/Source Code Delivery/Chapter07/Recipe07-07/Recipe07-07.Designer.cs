namespace Apress.VisualCSharpRecipes.Chapter07
{
    partial class Recipe07_07
    {
        private System.Windows.Forms.ListBox listBox1;
        private System.Windows.Forms.Button cmdTest;
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
            this.listBox1 = new System.Windows.Forms.ListBox();
            this.cmdTest = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // listBox1
            // 
            this.listBox1.FormattingEnabled = true;
            this.listBox1.Location = new System.Drawing.Point(8, 8);
            this.listBox1.Name = "listBox1";
            this.listBox1.Size = new System.Drawing.Size(248, 160);
            this.listBox1.TabIndex = 0;
            // 
            // cmdTest
            // 
            this.cmdTest.Location = new System.Drawing.Point(12, 184);
            this.cmdTest.Name = "cmdTest";
            this.cmdTest.Size = new System.Drawing.Size(132, 28);
            this.cmdTest.TabIndex = 1;
            this.cmdTest.Text = "Test Scroll";
            this.cmdTest.Click += new System.EventHandler(this.cmdTest_Click);
            // 
            // Recipe07_07
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(292, 246);
            this.Controls.Add(this.cmdTest);
            this.Controls.Add(this.listBox1);
            this.Name = "Recipe07_07";
            this.Text = "Recipe07_07";
            this.ResumeLayout(false);

        }

        #endregion
    }
}