namespace Apress.VisualCSharpRecipes.Chapter07
{
    partial class Recipe07_17
    {
        private System.Windows.Forms.Button Button1;
        private System.Windows.Forms.ErrorProvider errProvider;
        private System.Windows.Forms.Label Label3;
        private System.Windows.Forms.TextBox txtEmail;
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
            this.Button1 = new System.Windows.Forms.Button();
            this.errProvider = new System.Windows.Forms.ErrorProvider(this.components);
            this.Label3 = new System.Windows.Forms.Label();
            this.txtEmail = new System.Windows.Forms.TextBox();
            ((System.ComponentModel.ISupportInitialize)(this.errProvider)).BeginInit();
            this.SuspendLayout();
            // 
            // Button1
            // 
            this.Button1.Location = new System.Drawing.Point(212, 80);
            this.Button1.Name = "Button1";
            this.Button1.Size = new System.Drawing.Size(76, 24);
            this.Button1.TabIndex = 12;
            this.Button1.Text = "OK";
            this.Button1.Click += new System.EventHandler(this.Button1_Click);
            // 
            // errProvider
            // 
            this.errProvider.ContainerControl = this;
            this.errProvider.DataMember = "";
            // 
            // Label3
            // 
            this.Label3.Location = new System.Drawing.Point(24, 24);
            this.Label3.Name = "Label3";
            this.Label3.Size = new System.Drawing.Size(40, 16);
            this.Label3.TabIndex = 14;
            this.Label3.Text = "Email:";
            // 
            // txtEmail
            // 
            this.txtEmail.Location = new System.Drawing.Point(68, 20);
            this.txtEmail.Name = "txtEmail";
            this.txtEmail.Size = new System.Drawing.Size(220, 21);
            this.txtEmail.TabIndex = 13;
            this.txtEmail.Leave += new System.EventHandler(this.txtEmail_Leave);
            // 
            // Recipe07_17
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(328, 126);
            this.Controls.Add(this.Label3);
            this.Controls.Add(this.txtEmail);
            this.Controls.Add(this.Button1);
            this.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Name = "Recipe07_17";
            this.Text = "Recipe07_17";
            ((System.ComponentModel.ISupportInitialize)(this.errProvider)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
    }
}