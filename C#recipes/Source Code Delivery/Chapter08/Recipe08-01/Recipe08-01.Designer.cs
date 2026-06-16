namespace Apress.VisualCSharpRecipes.Chapter08
{
    partial class Recipe08_01
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
            this.pnlFonts = new System.Windows.Forms.Panel();
            this.SuspendLayout();
            // 
            // pnlFonts
            // 
            this.pnlFonts.AutoScroll = true;
            this.pnlFonts.Dock = System.Windows.Forms.DockStyle.Fill;
            this.pnlFonts.Location = new System.Drawing.Point(0, 0);
            this.pnlFonts.Name = "pnlFonts";
            this.pnlFonts.Size = new System.Drawing.Size(299, 276);
            this.pnlFonts.TabIndex = 0;
            // 
            // Recipe08_01
            // 
            this.ClientSize = new System.Drawing.Size(299, 276);
            this.Controls.Add(this.pnlFonts);
            this.Name = "Recipe08_01";
            this.Text = "List of Installed Fonts";
            this.Load += new System.EventHandler(this.Recipe08_01_Load);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel pnlFonts;
    }
}

