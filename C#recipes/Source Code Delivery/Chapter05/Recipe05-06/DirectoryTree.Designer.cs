namespace Apress.VisualCSharpRecipes.Chapter05
{
    partial class DirectoryTree
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
            this.treeDirectory = new System.Windows.Forms.TreeView();
            this.SuspendLayout();
            // 
            // treeDirectory
            // 
            this.treeDirectory.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                        | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.treeDirectory.Location = new System.Drawing.Point(12, 12);
            this.treeDirectory.Name = "treeDirectory";
            this.treeDirectory.Size = new System.Drawing.Size(204, 268);
            this.treeDirectory.TabIndex = 0;
            this.treeDirectory.BeforeExpand += new System.Windows.Forms.TreeViewCancelEventHandler(this.treeDirectory_BeforeExpand);
            // 
            // DirectoryTree
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(228, 292);
            this.Controls.Add(this.treeDirectory);
            this.Name = "DirectoryTree";
            this.Text = "Directory Tree";
            this.Load += new System.EventHandler(this.DirectoryTree_Load);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TreeView treeDirectory;
    }
}

