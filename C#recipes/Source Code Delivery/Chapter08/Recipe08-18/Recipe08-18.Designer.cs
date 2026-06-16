namespace Apress.VisualCSharpRecipes.Chapter08
{
    partial class Recipe08_18
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
            this.cmdRefresh = new System.Windows.Forms.Button();
            this.lstJobs = new System.Windows.Forms.ListBox();
            this.cmdPause = new System.Windows.Forms.Button();
            this.cmdResume = new System.Windows.Forms.Button();
            this.txtJobInfo = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // cmdRefresh
            // 
            this.cmdRefresh.Location = new System.Drawing.Point(12, 51);
            this.cmdRefresh.Name = "cmdRefresh";
            this.cmdRefresh.Size = new System.Drawing.Size(75, 23);
            this.cmdRefresh.TabIndex = 0;
            this.cmdRefresh.Text = "Refresh";
            this.cmdRefresh.UseVisualStyleBackColor = true;
            this.cmdRefresh.Click += new System.EventHandler(this.cmdRefresh_Click);
            // 
            // lstJobs
            // 
            this.lstJobs.FormattingEnabled = true;
            this.lstJobs.Location = new System.Drawing.Point(93, 51);
            this.lstJobs.Name = "lstJobs";
            this.lstJobs.Size = new System.Drawing.Size(187, 199);
            this.lstJobs.TabIndex = 1;
            this.lstJobs.SelectedIndexChanged += new System.EventHandler(this.lstJobs_SelectedIndexChanged);
            // 
            // cmdPause
            // 
            this.cmdPause.Location = new System.Drawing.Point(12, 80);
            this.cmdPause.Name = "cmdPause";
            this.cmdPause.Size = new System.Drawing.Size(75, 23);
            this.cmdPause.TabIndex = 0;
            this.cmdPause.Text = "Pause";
            this.cmdPause.UseVisualStyleBackColor = true;
            this.cmdPause.Click += new System.EventHandler(this.cmdPause_Click);
            // 
            // cmdResume
            // 
            this.cmdResume.Location = new System.Drawing.Point(12, 109);
            this.cmdResume.Name = "cmdResume";
            this.cmdResume.Size = new System.Drawing.Size(75, 23);
            this.cmdResume.TabIndex = 0;
            this.cmdResume.Text = "Resume";
            this.cmdResume.UseVisualStyleBackColor = true;
            this.cmdResume.Click += new System.EventHandler(this.cmdResume_Click);
            // 
            // txtJobInfo
            // 
            this.txtJobInfo.Location = new System.Drawing.Point(12, 25);
            this.txtJobInfo.Name = "txtJobInfo";
            this.txtJobInfo.ReadOnly = true;
            this.txtJobInfo.Size = new System.Drawing.Size(268, 20);
            this.txtJobInfo.TabIndex = 2;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 9);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(58, 13);
            this.label1.TabIndex = 3;
            this.label1.Text = "Job Name:";
            // 
            // Recipe08_18
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(296, 269);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.txtJobInfo);
            this.Controls.Add(this.lstJobs);
            this.Controls.Add(this.cmdResume);
            this.Controls.Add(this.cmdPause);
            this.Controls.Add(this.cmdRefresh);
            this.Name = "Recipe08_18";
            this.Text = "Print Queue Test ";
            this.Load += new System.EventHandler(this.Recipe08_18_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button cmdRefresh;
        private System.Windows.Forms.ListBox lstJobs;
        private System.Windows.Forms.Button cmdPause;
        private System.Windows.Forms.Button cmdResume;
        private System.Windows.Forms.TextBox txtJobInfo;
        private System.Windows.Forms.Label label1;
    }
}

