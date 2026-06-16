namespace Apress.VisualCSharpRecipes.Chapter08
{
    partial class Recipe08_07
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
            this.components = new System.ComponentModel.Container();
            this.tmrRefresh = new System.Windows.Forms.Timer(this.components);
            this.chkUseDoubleBuffering = new System.Windows.Forms.CheckBox();
            this.SuspendLayout();
            // 
            // tmrRefresh
            // 
            this.tmrRefresh.Tick += new System.EventHandler(this.tmrRefresh_Tick);
            // 
            // chkUseDoubleBuffering
            // 
            this.chkUseDoubleBuffering.AutoSize = true;
            this.chkUseDoubleBuffering.Location = new System.Drawing.Point(12, 3);
            this.chkUseDoubleBuffering.Name = "chkUseDoubleBuffering";
            this.chkUseDoubleBuffering.Size = new System.Drawing.Size(127, 17);
            this.chkUseDoubleBuffering.TabIndex = 0;
            this.chkUseDoubleBuffering.Text = "Use Double Buffering";
            this.chkUseDoubleBuffering.UseVisualStyleBackColor = true;
            this.chkUseDoubleBuffering.CheckedChanged += new System.EventHandler(this.chkUseDoubleBuffering_CheckedChanged);
            // 
            // Recipe08_07
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(299, 277);
            this.Controls.Add(this.chkUseDoubleBuffering);
            this.Name = "Recipe08_07";
            this.Text = "Double Buffering";
            this.Paint += new System.Windows.Forms.PaintEventHandler(this.Recipe08_07_Paint);
            this.Load += new System.EventHandler(this.Recipe08_07_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Timer tmrRefresh;
        private System.Windows.Forms.CheckBox chkUseDoubleBuffering;
    }
}

