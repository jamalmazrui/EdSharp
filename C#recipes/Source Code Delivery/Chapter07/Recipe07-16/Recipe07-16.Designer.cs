namespace Apress.VisualCSharpRecipes.Chapter07
{
    partial class Recipe07_16
    {
        private System.Windows.Forms.NotifyIcon notifyIcon;
        private System.Timers.Timer timer;
        
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
            this.notifyIcon = new System.Windows.Forms.NotifyIcon();
            this.timer = new System.Timers.Timer();
            this.SuspendLayout();
            // 
            // notifyIcon
            // 
            this.notifyIcon.Text = "Recipe07-16";
            this.notifyIcon.Visible = true;
            // 
            // timer
            // 
            this.timer.Enabled = true;
            this.timer.Interval = 1000D;
            this.timer.SynchronizingObject = this;
            this.timer.Elapsed += new System.Timers.ElapsedEventHandler(this.timer_Elapsed);
            // 
            // Recipe07_16
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(340, 174);
            this.Name = "Recipe07_16";
            this.Text = "Recipe07_16";
            this.ResumeLayout(false);

        }

        #endregion
    }
}