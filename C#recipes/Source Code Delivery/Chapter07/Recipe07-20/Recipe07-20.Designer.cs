namespace Apress.VisualCSharpRecipes.Chapter07
{
    partial class Recipe07_20
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
            this.webBrowser1 = new System.Windows.Forms.WebBrowser();
            this.goButton = new System.Windows.Forms.Button();
            this.textURL = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.backButton = new System.Windows.Forms.Button();
            this.homeButton = new System.Windows.Forms.Button();
            this.forwarButton = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // webBrowser1
            // 
            this.webBrowser1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.webBrowser1.Location = new System.Drawing.Point(-2, 2);
            this.webBrowser1.Name = "webBrowser1";
            this.webBrowser1.Size = new System.Drawing.Size(685, 190);
            this.webBrowser1.TabIndex = 3;
            this.webBrowser1.DocumentCompleted += new System.Windows.Forms.WebBrowserDocumentCompletedEventHandler(this.webBrowser1_DocumentCompleted);
            // 
            // goButton
            // 
            this.goButton.Location = new System.Drawing.Point(435, 216);
            this.goButton.Name = "goButton";
            this.goButton.Size = new System.Drawing.Size(48, 23);
            this.goButton.TabIndex = 1;
            this.goButton.Text = "Go";
            this.goButton.Click += new System.EventHandler(this.goButton_Click);
            // 
            // textURL
            // 
            this.textURL.Location = new System.Drawing.Point(240, 217);
            this.textURL.Name = "textURL";
            this.textURL.Size = new System.Drawing.Size(189, 20);
            this.textURL.TabIndex = 2;
            this.textURL.Text = "http://www.apress.com";
            // 
            // label1
            // 
            this.label1.Location = new System.Drawing.Point(206, 221);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(31, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "Go to:";
            // 
            // backButton
            // 
            this.backButton.Enabled = false;
            this.backButton.Location = new System.Drawing.Point(227, 249);
            this.backButton.Name = "backButton";
            this.backButton.Size = new System.Drawing.Size(75, 23);
            this.backButton.TabIndex = 0;
            this.backButton.Text = "<< Back";
            this.backButton.Click += new System.EventHandler(this.backButton_Click);
            // 
            // homeButton
            // 
            this.homeButton.Location = new System.Drawing.Point(308, 249);
            this.homeButton.Name = "homeButton";
            this.homeButton.Size = new System.Drawing.Size(75, 23);
            this.homeButton.TabIndex = 0;
            this.homeButton.Text = "Home";
            this.homeButton.Click += new System.EventHandler(this.homeButton_Click);
            // 
            // forwarButton
            // 
            this.forwarButton.Enabled = false;
            this.forwarButton.Location = new System.Drawing.Point(389, 249);
            this.forwarButton.Name = "forwarButton";
            this.forwarButton.Size = new System.Drawing.Size(75, 23);
            this.forwarButton.TabIndex = 0;
            this.forwarButton.Text = "Forward >>";
            this.forwarButton.Click += new System.EventHandler(this.forwarButton_Click);
            // 
            // Recipe07_20
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(684, 303);
            this.Controls.Add(this.forwarButton);
            this.Controls.Add(this.homeButton);
            this.Controls.Add(this.backButton);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.textURL);
            this.Controls.Add(this.goButton);
            this.Controls.Add(this.webBrowser1);
            this.Name = "Recipe07_20";
            this.Text = "Recipe07_20";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.WebBrowser webBrowser1;
        private System.Windows.Forms.Button goButton;
        private System.Windows.Forms.TextBox textURL;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button backButton;
        private System.Windows.Forms.Button homeButton;
        private System.Windows.Forms.Button forwarButton;

    }
}