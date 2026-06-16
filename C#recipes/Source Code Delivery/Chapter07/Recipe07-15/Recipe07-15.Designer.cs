namespace Apress.VisualCSharpRecipes.Chapter07
{
    partial class Recipe07_15
    {
        private System.Windows.Forms.Button cmdClose;
        private System.Windows.Forms.Label lblDrag;
        private System.Windows.Forms.Label Label1;        
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
            this.cmdClose = new System.Windows.Forms.Button();
            this.lblDrag = new System.Windows.Forms.Label();
            this.Label1 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // cmdClose
            // 
            this.cmdClose.Location = new System.Drawing.Point(108, 215);
            this.cmdClose.Name = "cmdClose";
            this.cmdClose.Size = new System.Drawing.Size(76, 20);
            this.cmdClose.TabIndex = 5;
            this.cmdClose.Text = "Exit";
            this.cmdClose.Click += new System.EventHandler(this.cmdClose_Click);
            // 
            // lblDrag
            // 
            this.lblDrag.BackColor = System.Drawing.Color.Navy;
            this.lblDrag.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblDrag.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblDrag.ForeColor = System.Drawing.Color.White;
            this.lblDrag.Location = new System.Drawing.Point(98, 150);
            this.lblDrag.Name = "lblDrag";
            this.lblDrag.Size = new System.Drawing.Size(96, 44);
            this.lblDrag.TabIndex = 4;
            this.lblDrag.Text = "Click here and drag to move the form!";
            this.lblDrag.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.lblDrag.MouseDown += new System.Windows.Forms.MouseEventHandler(this.lblDrag_MouseDown);
            this.lblDrag.MouseMove += new System.Windows.Forms.MouseEventHandler(this.lblDrag_MouseMove);
            this.lblDrag.MouseUp += new System.Windows.Forms.MouseEventHandler(this.lblDrag_MouseUp);
            // 
            // Label1
            // 
            this.Label1.Location = new System.Drawing.Point(28, 31);
            this.Label1.Name = "Label1";
            this.Label1.Size = new System.Drawing.Size(230, 100);
            this.Label1.TabIndex = 3;
            this.Label1.Text = "This form has a control-style border, but no \r\nminimize box, maximize box, contro" +
                "l box, or \r\ntext caption.";
            // 
            // Recipe07_15
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(292, 266);
            this.ControlBox = false;
            this.Controls.Add(this.cmdClose);
            this.Controls.Add(this.lblDrag);
            this.Controls.Add(this.Label1);
            this.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "Recipe07_15";
            this.ResumeLayout(false);

        }

        #endregion
    }
}