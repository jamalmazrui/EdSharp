namespace Apress.VisualCSharpRecipes.Chapter07
{
    partial class Recipe07_05Child
    {
        private System.Windows.Forms.Button cmdShowAllWindows;
        private System.Windows.Forms.Label label;
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
            this.cmdShowAllWindows = new System.Windows.Forms.Button();
            this.label = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // cmdShowAllWindows
            // 
            this.cmdShowAllWindows.Location = new System.Drawing.Point(84, 64);
            this.cmdShowAllWindows.Name = "cmdShowAllWindows";
            this.cmdShowAllWindows.Size = new System.Drawing.Size(104, 36);
            this.cmdShowAllWindows.TabIndex = 0;
            this.cmdShowAllWindows.Text = "Show Values From All Windows";
            this.cmdShowAllWindows.Click += new System.EventHandler(this.cmdShowAllWindows_Click);
            // 
            // label
            // 
            this.label.Location = new System.Drawing.Point(28, 16);
            this.label.Name = "label";
            this.label.Size = new System.Drawing.Size(232, 24);
            this.label.TabIndex = 1;
            this.label.Text = "label";
            // 
            // Recipe07_05Child
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(284, 122);
            this.Controls.Add(this.label);
            this.Controls.Add(this.cmdShowAllWindows);
            this.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Name = "Recipe07_05Child";
            this.Text = "Recipe07_05Child";
            this.ResumeLayout(false);

        }

        #endregion
    }
}