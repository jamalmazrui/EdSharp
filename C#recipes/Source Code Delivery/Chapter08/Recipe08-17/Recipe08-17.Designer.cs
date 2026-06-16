namespace Apress.VisualCSharpRecipes.Chapter08
{
    partial class Recipe08_17
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
            this.cmdPrint = new System.Windows.Forms.Button();
            this.lstZoom = new System.Windows.Forms.ListBox();
            this.printPreviewControl = new System.Windows.Forms.PrintPreviewControl();
            this.SuspendLayout();
            // 
            // cmdPrint
            // 
            this.cmdPrint.Location = new System.Drawing.Point(135, 231);
            this.cmdPrint.Name = "cmdPrint";
            this.cmdPrint.Size = new System.Drawing.Size(95, 23);
            this.cmdPrint.TabIndex = 2;
            this.cmdPrint.Text = "Print Preview";
            this.cmdPrint.UseVisualStyleBackColor = true;
            this.cmdPrint.Click += new System.EventHandler(this.cmdPrint_Click);
            // 
            // lstZoom
            // 
            this.lstZoom.FormattingEnabled = true;
            this.lstZoom.Location = new System.Drawing.Point(264, 12);
            this.lstZoom.Name = "lstZoom";
            this.lstZoom.Size = new System.Drawing.Size(89, 212);
            this.lstZoom.TabIndex = 3;
            // 
            // printPreviewControl
            // 
            this.printPreviewControl.Location = new System.Drawing.Point(12, 12);
            this.printPreviewControl.Name = "printPreviewControl";
            this.printPreviewControl.Size = new System.Drawing.Size(246, 213);
            this.printPreviewControl.TabIndex = 4;
            // 
            // Recipe08_17
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(368, 270);
            this.Controls.Add(this.printPreviewControl);
            this.Controls.Add(this.lstZoom);
            this.Controls.Add(this.cmdPrint);
            this.Name = "Recipe08_17";
            this.Text = "Print Preview ";
            this.Load += new System.EventHandler(this.Recipe08_17_Load);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button cmdPrint;
        private System.Windows.Forms.ListBox lstZoom;
        private System.Windows.Forms.PrintPreviewControl printPreviewControl;
    }
}

