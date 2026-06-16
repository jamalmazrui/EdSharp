namespace Apress.VisualCSharpRecipes.Chapter07
{
    partial class Recipe07_04
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
            this.listForms = new System.Windows.Forms.ListBox();
            this.label1 = new System.Windows.Forms.Label();
            this.txtFormName = new System.Windows.Forms.TextBox();
            this.btnNewForm = new System.Windows.Forms.Button();
            this.btnRefresh = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // listForms
            // 
            this.listForms.FormattingEnabled = true;
            this.listForms.Location = new System.Drawing.Point(12, 28);
            this.listForms.Name = "listForms";
            this.listForms.Size = new System.Drawing.Size(268, 160);
            this.listForms.TabIndex = 0;
            this.listForms.SelectedIndexChanged += new System.EventHandler(this.listForms_SelectedIndexChanged);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 12);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(38, 13);
            this.label1.TabIndex = 1;
            this.label1.Text = "Forms:";
            // 
            // txtFormName
            // 
            this.txtFormName.Location = new System.Drawing.Point(12, 194);
            this.txtFormName.Name = "txtFormName";
            this.txtFormName.Size = new System.Drawing.Size(170, 20);
            this.txtFormName.TabIndex = 2;
            // 
            // btnNewForm
            // 
            this.btnNewForm.Location = new System.Drawing.Point(200, 194);
            this.btnNewForm.Name = "btnNewForm";
            this.btnNewForm.Size = new System.Drawing.Size(79, 19);
            this.btnNewForm.TabIndex = 3;
            this.btnNewForm.Text = "&New Form";
            this.btnNewForm.UseVisualStyleBackColor = true;
            this.btnNewForm.Click += new System.EventHandler(this.btnNewForm_Click);
            // 
            // btnRefresh
            // 
            this.btnRefresh.Location = new System.Drawing.Point(97, 230);
            this.btnRefresh.Name = "btnRefresh";
            this.btnRefresh.Size = new System.Drawing.Size(95, 24);
            this.btnRefresh.TabIndex = 4;
            this.btnRefresh.Text = "&Refresh List";
            this.btnRefresh.UseVisualStyleBackColor = true;
            this.btnRefresh.Click += new System.EventHandler(this.btnRefresh_Click);
            // 
            // Recipe07_04
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(292, 266);
            this.Controls.Add(this.btnRefresh);
            this.Controls.Add(this.btnNewForm);
            this.Controls.Add(this.txtFormName);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.listForms);
            this.Name = "Recipe07_04";
            this.Text = "Recipe07_04";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ListBox listForms;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox txtFormName;
        private System.Windows.Forms.Button btnNewForm;
        private System.Windows.Forms.Button btnRefresh;
    }
}