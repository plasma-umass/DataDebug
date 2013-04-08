namespace ErrorClassifier
{
    partial class Form1
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
            this.decimalOmission = new System.Windows.Forms.Button();
            this.digitOmission = new System.Windows.Forms.Button();
            this.digitRepeat = new System.Windows.Forms.Button();
            this.decimalPoint = new System.Windows.Forms.Button();
            this.signOmission = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.enteredText = new System.Windows.Forms.TextBox();
            this.originalText = new System.Windows.Forms.TextBox();
            this.wrongDigit = new System.Windows.Forms.Button();
            this.extraDigit = new System.Windows.Forms.Button();
            this.digitTransposition = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // decimalOmission
            // 
            this.decimalOmission.Location = new System.Drawing.Point(6, 116);
            this.decimalOmission.Name = "decimalOmission";
            this.decimalOmission.Size = new System.Drawing.Size(130, 23);
            this.decimalOmission.TabIndex = 17;
            this.decimalOmission.Text = "Test decimal omission";
            this.decimalOmission.UseVisualStyleBackColor = true;
            this.decimalOmission.Click += new System.EventHandler(this.decimalOmission_Click);
            // 
            // digitOmission
            // 
            this.digitOmission.Location = new System.Drawing.Point(153, 87);
            this.digitOmission.Name = "digitOmission";
            this.digitOmission.Size = new System.Drawing.Size(130, 23);
            this.digitOmission.TabIndex = 16;
            this.digitOmission.Text = "Test digit omission";
            this.digitOmission.UseVisualStyleBackColor = true;
            this.digitOmission.Click += new System.EventHandler(this.digitOmission_Click);
            // 
            // digitRepeat
            // 
            this.digitRepeat.Location = new System.Drawing.Point(6, 87);
            this.digitRepeat.Name = "digitRepeat";
            this.digitRepeat.Size = new System.Drawing.Size(130, 23);
            this.digitRepeat.TabIndex = 15;
            this.digitRepeat.Text = "Test digit repeat";
            this.digitRepeat.UseVisualStyleBackColor = true;
            this.digitRepeat.Click += new System.EventHandler(this.digitRepeat_Click);
            // 
            // decimalPoint
            // 
            this.decimalPoint.Location = new System.Drawing.Point(153, 58);
            this.decimalPoint.Name = "decimalPoint";
            this.decimalPoint.Size = new System.Drawing.Size(130, 23);
            this.decimalPoint.TabIndex = 14;
            this.decimalPoint.Text = "Test misplaced decimal";
            this.decimalPoint.UseVisualStyleBackColor = true;
            this.decimalPoint.Click += new System.EventHandler(this.decimalPoint_Click);
            // 
            // signOmission
            // 
            this.signOmission.Location = new System.Drawing.Point(6, 58);
            this.signOmission.Name = "signOmission";
            this.signOmission.Size = new System.Drawing.Size(130, 23);
            this.signOmission.TabIndex = 13;
            this.signOmission.Text = "Test sign omission";
            this.signOmission.UseVisualStyleBackColor = true;
            this.signOmission.Click += new System.EventHandler(this.signOmission_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(150, 6);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(76, 13);
            this.label2.TabIndex = 12;
            this.label2.Text = "Entered value:";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(3, 6);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(74, 13);
            this.label1.TabIndex = 11;
            this.label1.Text = "Original value:";
            // 
            // enteredText
            // 
            this.enteredText.Location = new System.Drawing.Point(153, 24);
            this.enteredText.Name = "enteredText";
            this.enteredText.Size = new System.Drawing.Size(130, 20);
            this.enteredText.TabIndex = 10;
            // 
            // originalText
            // 
            this.originalText.Location = new System.Drawing.Point(6, 24);
            this.originalText.Name = "originalText";
            this.originalText.Size = new System.Drawing.Size(130, 20);
            this.originalText.TabIndex = 9;
            // 
            // wrongDigit
            // 
            this.wrongDigit.Location = new System.Drawing.Point(153, 116);
            this.wrongDigit.Name = "wrongDigit";
            this.wrongDigit.Size = new System.Drawing.Size(130, 23);
            this.wrongDigit.TabIndex = 18;
            this.wrongDigit.Text = "Test wrong digit";
            this.wrongDigit.UseVisualStyleBackColor = true;
            this.wrongDigit.Click += new System.EventHandler(this.wrongDigit_Click);
            // 
            // extraDigit
            // 
            this.extraDigit.Location = new System.Drawing.Point(6, 145);
            this.extraDigit.Name = "extraDigit";
            this.extraDigit.Size = new System.Drawing.Size(130, 23);
            this.extraDigit.TabIndex = 19;
            this.extraDigit.Text = "Test extra digit";
            this.extraDigit.UseVisualStyleBackColor = true;
            this.extraDigit.Click += new System.EventHandler(this.extraDigit_Click);
            // 
            // digitTransposition
            // 
            this.digitTransposition.Location = new System.Drawing.Point(153, 145);
            this.digitTransposition.Name = "digitTransposition";
            this.digitTransposition.Size = new System.Drawing.Size(130, 23);
            this.digitTransposition.TabIndex = 20;
            this.digitTransposition.Text = "Test digit transposition";
            this.digitTransposition.UseVisualStyleBackColor = true;
            this.digitTransposition.Click += new System.EventHandler(this.digitTransposition_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(293, 176);
            this.Controls.Add(this.digitTransposition);
            this.Controls.Add(this.extraDigit);
            this.Controls.Add(this.wrongDigit);
            this.Controls.Add(this.decimalOmission);
            this.Controls.Add(this.digitOmission);
            this.Controls.Add(this.digitRepeat);
            this.Controls.Add(this.decimalPoint);
            this.Controls.Add(this.signOmission);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.enteredText);
            this.Controls.Add(this.originalText);
            this.Name = "Form1";
            this.Text = "Form1";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button decimalOmission;
        private System.Windows.Forms.Button digitOmission;
        private System.Windows.Forms.Button digitRepeat;
        private System.Windows.Forms.Button decimalPoint;
        private System.Windows.Forms.Button signOmission;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox enteredText;
        private System.Windows.Forms.TextBox originalText;
        private System.Windows.Forms.Button wrongDigit;
        private System.Windows.Forms.Button extraDigit;
        private System.Windows.Forms.Button digitTransposition;
    }
}