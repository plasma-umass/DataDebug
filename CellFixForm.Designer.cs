namespace DataDebug
{
    partial class CellFixForm
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
            this.FixText = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.AcceptFix = new System.Windows.Forms.Button();
            this.CancelFix = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // FixText
            // 
            this.FixText.Location = new System.Drawing.Point(12, 38);
            this.FixText.Name = "FixText";
            this.FixText.Size = new System.Drawing.Size(348, 20);
            this.FixText.TabIndex = 0;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(9, 9);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(130, 13);
            this.label1.TabIndex = 1;
            this.label1.Text = "Enter the corrected value:";
            // 
            // AcceptFix
            // 
            this.AcceptFix.Location = new System.Drawing.Point(285, 83);
            this.AcceptFix.Name = "AcceptFix";
            this.AcceptFix.Size = new System.Drawing.Size(75, 23);
            this.AcceptFix.TabIndex = 2;
            this.AcceptFix.Text = "Fix";
            this.AcceptFix.UseVisualStyleBackColor = true;
            this.AcceptFix.Click += new System.EventHandler(this.AcceptFix_Click);
            // 
            // CancelFix
            // 
            this.CancelFix.Location = new System.Drawing.Point(204, 83);
            this.CancelFix.Name = "CancelFix";
            this.CancelFix.Size = new System.Drawing.Size(75, 23);
            this.CancelFix.TabIndex = 3;
            this.CancelFix.Text = "Cancel";
            this.CancelFix.UseVisualStyleBackColor = true;
            this.CancelFix.Click += new System.EventHandler(this.CancelFix_Click);
            // 
            // CellFixForm
            // 
            this.AcceptButton = this.AcceptFix;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(372, 125);
            this.Controls.Add(this.CancelFix);
            this.Controls.Add(this.AcceptFix);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.FixText);
            this.Name = "CellFixForm";
            this.Text = "CellFixForm";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox FixText;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button AcceptFix;
        private System.Windows.Forms.Button CancelFix;
    }
}