namespace Apress.VisualCSharpRecipes.Chapter07
{
    partial class Recipe07_08
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
            this.mskTextBox = new System.Windows.Forms.MaskedTextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.lblActiveMask = new System.Windows.Forms.Label();
            this.btnTime = new System.Windows.Forms.Button();
            this.btnUSZip = new System.Windows.Forms.Button();
            this.btnCurrency = new System.Windows.Forms.Button();
            this.btnUKPost = new System.Windows.Forms.Button();
            this.btnDate = new System.Windows.Forms.Button();
            this.btnSecret = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // mskTextBox
            // 
            this.mskTextBox.BeepOnError = true;
            this.mskTextBox.Location = new System.Drawing.Point(15, 41);
            this.mskTextBox.Name = "mskTextBox";
            this.mskTextBox.Size = new System.Drawing.Size(265, 20);
            this.mskTextBox.TabIndex = 1;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 23);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(94, 13);
            this.label1.TabIndex = 2;
            this.label1.Text = "Active input mask:";
            // 
            // lblActiveMask
            // 
            this.lblActiveMask.AutoSize = true;
            this.lblActiveMask.Location = new System.Drawing.Point(112, 23);
            this.lblActiveMask.Name = "lblActiveMask";
            this.lblActiveMask.Size = new System.Drawing.Size(51, 13);
            this.lblActiveMask.TabIndex = 3;
            this.lblActiveMask.Text = "Any input";
            // 
            // btnTime
            // 
            this.btnTime.Location = new System.Drawing.Point(15, 85);
            this.btnTime.Name = "btnTime";
            this.btnTime.Size = new System.Drawing.Size(83, 23);
            this.btnTime.TabIndex = 4;
            this.btnTime.Text = "24-hour time";
            this.btnTime.UseVisualStyleBackColor = true;
            this.btnTime.Click += new System.EventHandler(this.btnTime_Click);
            // 
            // btnUSZip
            // 
            this.btnUSZip.Location = new System.Drawing.Point(15, 130);
            this.btnUSZip.Name = "btnUSZip";
            this.btnUSZip.Size = new System.Drawing.Size(83, 23);
            this.btnUSZip.TabIndex = 5;
            this.btnUSZip.Text = "US ZIP";
            this.btnUSZip.UseVisualStyleBackColor = true;
            this.btnUSZip.Click += new System.EventHandler(this.btnUSZip_Click);
            // 
            // btnCurrency
            // 
            this.btnCurrency.Location = new System.Drawing.Point(104, 85);
            this.btnCurrency.Name = "btnCurrency";
            this.btnCurrency.Size = new System.Drawing.Size(87, 23);
            this.btnCurrency.TabIndex = 7;
            this.btnCurrency.Text = "Currency";
            this.btnCurrency.UseVisualStyleBackColor = true;
            this.btnCurrency.Click += new System.EventHandler(this.btnCurrency_Click);
            // 
            // btnUKPost
            // 
            this.btnUKPost.Location = new System.Drawing.Point(104, 130);
            this.btnUKPost.Name = "btnUKPost";
            this.btnUKPost.Size = new System.Drawing.Size(87, 23);
            this.btnUKPost.TabIndex = 8;
            this.btnUKPost.Text = "UK Postcode";
            this.btnUKPost.UseVisualStyleBackColor = true;
            this.btnUKPost.Click += new System.EventHandler(this.btnUKPost_Click);
            // 
            // btnDate
            // 
            this.btnDate.Location = new System.Drawing.Point(197, 85);
            this.btnDate.Name = "btnDate";
            this.btnDate.Size = new System.Drawing.Size(83, 23);
            this.btnDate.TabIndex = 10;
            this.btnDate.Text = "Short Date";
            this.btnDate.UseVisualStyleBackColor = true;
            this.btnDate.Click += new System.EventHandler(this.btnDate_Click);
            // 
            // btnSecret
            // 
            this.btnSecret.Location = new System.Drawing.Point(197, 130);
            this.btnSecret.Name = "btnSecret";
            this.btnSecret.Size = new System.Drawing.Size(83, 23);
            this.btnSecret.TabIndex = 11;
            this.btnSecret.Text = "Secret PIN";
            this.btnSecret.UseVisualStyleBackColor = true;
            this.btnSecret.Click += new System.EventHandler(this.btnSecret_Click);
            // 
            // Recipe07_08
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(292, 196);
            this.Controls.Add(this.btnSecret);
            this.Controls.Add(this.btnDate);
            this.Controls.Add(this.btnUKPost);
            this.Controls.Add(this.btnCurrency);
            this.Controls.Add(this.btnUSZip);
            this.Controls.Add(this.btnTime);
            this.Controls.Add(this.lblActiveMask);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.mskTextBox);
            this.Name = "Recipe07_08";
            this.Text = "Recipe07_08";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.MaskedTextBox mskTextBox;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label lblActiveMask;
        private System.Windows.Forms.Button btnTime;
        private System.Windows.Forms.Button btnUSZip;
        private System.Windows.Forms.Button btnCurrency;
        private System.Windows.Forms.Button btnUKPost;
        private System.Windows.Forms.Button btnDate;
        private System.Windows.Forms.Button btnSecret;
    }
}