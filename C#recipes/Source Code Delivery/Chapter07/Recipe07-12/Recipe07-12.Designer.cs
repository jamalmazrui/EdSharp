namespace Apress.VisualCSharpRecipes.Chapter07
{
    partial class Recipe07_12
    {
        private System.Windows.Forms.MainMenu MainMenu1;
        private System.Windows.Forms.MenuItem mnuFile;
        private System.Windows.Forms.MenuItem mnuOpen;
        private System.Windows.Forms.MenuItem mnuSave;
        private System.Windows.Forms.MenuItem mnuExit;
        private System.Windows.Forms.MenuItem MenuItem5;
        private System.Windows.Forms.MenuItem MenuItem6;
        private System.Windows.Forms.MenuItem MenuItem7;
        private System.Windows.Forms.MenuItem MenuItem8;
        private System.Windows.Forms.TextBox TextBox1;
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
            this.components = new System.ComponentModel.Container();
            this.MainMenu1 = new System.Windows.Forms.MainMenu(this.components);
            this.mnuFile = new System.Windows.Forms.MenuItem();
            this.mnuOpen = new System.Windows.Forms.MenuItem();
            this.mnuSave = new System.Windows.Forms.MenuItem();
            this.mnuExit = new System.Windows.Forms.MenuItem();
            this.MenuItem5 = new System.Windows.Forms.MenuItem();
            this.MenuItem6 = new System.Windows.Forms.MenuItem();
            this.MenuItem7 = new System.Windows.Forms.MenuItem();
            this.MenuItem8 = new System.Windows.Forms.MenuItem();
            this.TextBox1 = new System.Windows.Forms.TextBox();
            this.SuspendLayout();
            // 
            // MainMenu1
            // 
            this.MainMenu1.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
            this.mnuFile,
            this.MenuItem5});
            // 
            // mnuFile
            // 
            this.mnuFile.Index = 0;
            this.mnuFile.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
            this.mnuOpen,
            this.mnuSave,
            this.mnuExit});
            this.mnuFile.Text = "File";
            // 
            // mnuOpen
            // 
            this.mnuOpen.Index = 0;
            this.mnuOpen.Text = "Open";
            this.mnuOpen.Click += new System.EventHandler(this.mnuOpen_Click);
            // 
            // mnuSave
            // 
            this.mnuSave.Index = 1;
            this.mnuSave.Text = "Save";
            this.mnuSave.Click += new System.EventHandler(this.mnuSave_Click);
            // 
            // mnuExit
            // 
            this.mnuExit.Index = 2;
            this.mnuExit.Text = "Exit";
            this.mnuExit.Click += new System.EventHandler(this.mnuExit_Click);
            // 
            // MenuItem5
            // 
            this.MenuItem5.Index = 1;
            this.MenuItem5.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
            this.MenuItem6,
            this.MenuItem7,
            this.MenuItem8});
            this.MenuItem5.Text = "Edit";
            // 
            // MenuItem6
            // 
            this.MenuItem6.Index = 0;
            this.MenuItem6.Text = "Cut";
            // 
            // MenuItem7
            // 
            this.MenuItem7.Index = 1;
            this.MenuItem7.Text = "Copy";
            // 
            // MenuItem8
            // 
            this.MenuItem8.Index = 2;
            this.MenuItem8.Text = "Paste";
            // 
            // TextBox1
            // 
            this.TextBox1.Location = new System.Drawing.Point(44, 56);
            this.TextBox1.Multiline = true;
            this.TextBox1.Name = "TextBox1";
            this.TextBox1.Size = new System.Drawing.Size(180, 88);
            this.TextBox1.TabIndex = 1;
            this.TextBox1.Text = "Right click here to view the cloned context menu.";
            this.TextBox1.MouseDown += new System.Windows.Forms.MouseEventHandler(this.TextBox1_MouseDown);
            // 
            // Recipe07_12
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(292, 266);
            this.Controls.Add(this.TextBox1);
            this.Menu = this.MainMenu1;
            this.Name = "Recipe07_12";
            this.Text = "Recipe07_12";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
    }
}