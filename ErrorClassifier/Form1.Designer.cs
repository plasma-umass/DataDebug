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
            this.signError = new System.Windows.Forms.Button();
            this.signOmissionTextBox = new System.Windows.Forms.TextBox();
            this.digitRepeatTextBox = new System.Windows.Forms.TextBox();
            this.decimalOmissionTextBox = new System.Windows.Forms.TextBox();
            this.extraDigitTextBox = new System.Windows.Forms.TextBox();
            this.signErrorTextBox = new System.Windows.Forms.TextBox();
            this.misplacedDecimalTextBox = new System.Windows.Forms.TextBox();
            this.digitOmissionTextBox = new System.Windows.Forms.TextBox();
            this.wrongDigitTextBox = new System.Windows.Forms.TextBox();
            this.digitTranspositionTextBox = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.label8 = new System.Windows.Forms.Label();
            this.label9 = new System.Windows.Forms.Label();
            this.label10 = new System.Windows.Forms.Label();
            this.label11 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // decimalOmission
            // 
            this.decimalOmission.Font = new System.Drawing.Font("Lucida Console", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.decimalOmission.ForeColor = System.Drawing.Color.LightGray;
            this.decimalOmission.Location = new System.Drawing.Point(6, 116);
            this.decimalOmission.Name = "decimalOmission";
            this.decimalOmission.Size = new System.Drawing.Size(103, 23);
            this.decimalOmission.TabIndex = 17;
            this.decimalOmission.Text = "Decimal omission";
            this.decimalOmission.UseVisualStyleBackColor = true;
            this.decimalOmission.Visible = false;
            this.decimalOmission.Click += new System.EventHandler(this.decimalOmission_Click);
            // 
            // digitOmission
            // 
            this.digitOmission.Font = new System.Drawing.Font("Lucida Console", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.digitOmission.ForeColor = System.Drawing.Color.LightGray;
            this.digitOmission.Location = new System.Drawing.Point(177, 87);
            this.digitOmission.Name = "digitOmission";
            this.digitOmission.Size = new System.Drawing.Size(103, 23);
            this.digitOmission.TabIndex = 16;
            this.digitOmission.Text = "Test digit omission";
            this.digitOmission.UseVisualStyleBackColor = true;
            this.digitOmission.Visible = false;
            this.digitOmission.Click += new System.EventHandler(this.digitOmission_Click);
            // 
            // digitRepeat
            // 
            this.digitRepeat.Font = new System.Drawing.Font("Lucida Console", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.digitRepeat.ForeColor = System.Drawing.Color.LightGray;
            this.digitRepeat.Location = new System.Drawing.Point(6, 87);
            this.digitRepeat.Name = "digitRepeat";
            this.digitRepeat.Size = new System.Drawing.Size(103, 23);
            this.digitRepeat.TabIndex = 15;
            this.digitRepeat.Text = "Test digit repeat";
            this.digitRepeat.UseVisualStyleBackColor = true;
            this.digitRepeat.Visible = false;
            this.digitRepeat.Click += new System.EventHandler(this.digitRepeat_Click);
            // 
            // decimalPoint
            // 
            this.decimalPoint.Font = new System.Drawing.Font("Lucida Console", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.decimalPoint.ForeColor = System.Drawing.Color.LightGray;
            this.decimalPoint.Location = new System.Drawing.Point(177, 58);
            this.decimalPoint.Name = "decimalPoint";
            this.decimalPoint.Size = new System.Drawing.Size(103, 23);
            this.decimalPoint.TabIndex = 14;
            this.decimalPoint.Text = "Misplaced decimal";
            this.decimalPoint.UseVisualStyleBackColor = true;
            this.decimalPoint.Visible = false;
            this.decimalPoint.Click += new System.EventHandler(this.decimalPoint_Click);
            // 
            // signOmission
            // 
            this.signOmission.Font = new System.Drawing.Font("Lucida Console", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.signOmission.ForeColor = System.Drawing.Color.LightGray;
            this.signOmission.Location = new System.Drawing.Point(6, 58);
            this.signOmission.Name = "signOmission";
            this.signOmission.Size = new System.Drawing.Size(103, 23);
            this.signOmission.TabIndex = 13;
            this.signOmission.Text = "Test sign omission";
            this.signOmission.UseVisualStyleBackColor = true;
            this.signOmission.Visible = false;
            this.signOmission.Click += new System.EventHandler(this.signOmission_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Lucida Console", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.ForeColor = System.Drawing.Color.LightGray;
            this.label2.Location = new System.Drawing.Point(174, 6);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(103, 11);
            this.label2.TabIndex = 12;
            this.label2.Text = "Entered value:";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Lucida Console", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.ForeColor = System.Drawing.Color.LightGray;
            this.label1.Location = new System.Drawing.Point(3, 6);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(110, 11);
            this.label1.TabIndex = 11;
            this.label1.Text = "Original value:";
            // 
            // enteredText
            // 
            this.enteredText.BackColor = System.Drawing.Color.Black;
            this.enteredText.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.enteredText.Font = new System.Drawing.Font("Lucida Console", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.enteredText.ForeColor = System.Drawing.Color.LightGray;
            this.enteredText.Location = new System.Drawing.Point(177, 24);
            this.enteredText.Name = "enteredText";
            this.enteredText.Size = new System.Drawing.Size(168, 18);
            this.enteredText.TabIndex = 10;
            this.enteredText.TextChanged += new System.EventHandler(this.enteredText_TextChanged);
            // 
            // originalText
            // 
            this.originalText.BackColor = System.Drawing.Color.Black;
            this.originalText.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.originalText.Font = new System.Drawing.Font("Lucida Console", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.originalText.ForeColor = System.Drawing.Color.LightGray;
            this.originalText.Location = new System.Drawing.Point(6, 24);
            this.originalText.Name = "originalText";
            this.originalText.Size = new System.Drawing.Size(148, 18);
            this.originalText.TabIndex = 9;
            this.originalText.TextChanged += new System.EventHandler(this.originalText_TextChanged);
            // 
            // wrongDigit
            // 
            this.wrongDigit.Font = new System.Drawing.Font("Lucida Console", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.wrongDigit.ForeColor = System.Drawing.Color.LightGray;
            this.wrongDigit.Location = new System.Drawing.Point(177, 116);
            this.wrongDigit.Name = "wrongDigit";
            this.wrongDigit.Size = new System.Drawing.Size(103, 23);
            this.wrongDigit.TabIndex = 18;
            this.wrongDigit.Text = "Test wrong digit";
            this.wrongDigit.UseVisualStyleBackColor = true;
            this.wrongDigit.Visible = false;
            this.wrongDigit.Click += new System.EventHandler(this.wrongDigit_Click);
            // 
            // extraDigit
            // 
            this.extraDigit.Font = new System.Drawing.Font("Lucida Console", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.extraDigit.ForeColor = System.Drawing.Color.LightGray;
            this.extraDigit.Location = new System.Drawing.Point(6, 145);
            this.extraDigit.Name = "extraDigit";
            this.extraDigit.Size = new System.Drawing.Size(103, 23);
            this.extraDigit.TabIndex = 19;
            this.extraDigit.Text = "Test extra digit";
            this.extraDigit.UseVisualStyleBackColor = true;
            this.extraDigit.Visible = false;
            this.extraDigit.Click += new System.EventHandler(this.extraDigit_Click);
            // 
            // digitTransposition
            // 
            this.digitTransposition.Font = new System.Drawing.Font("Lucida Console", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.digitTransposition.ForeColor = System.Drawing.Color.LightGray;
            this.digitTransposition.Location = new System.Drawing.Point(177, 145);
            this.digitTransposition.Name = "digitTransposition";
            this.digitTransposition.Size = new System.Drawing.Size(103, 23);
            this.digitTransposition.TabIndex = 20;
            this.digitTransposition.Text = "Digit transposition";
            this.digitTransposition.UseVisualStyleBackColor = true;
            this.digitTransposition.Visible = false;
            this.digitTransposition.Click += new System.EventHandler(this.digitTransposition_Click);
            // 
            // signError
            // 
            this.signError.Font = new System.Drawing.Font("Lucida Console", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.signError.ForeColor = System.Drawing.Color.LightGray;
            this.signError.Location = new System.Drawing.Point(6, 174);
            this.signError.Name = "signError";
            this.signError.Size = new System.Drawing.Size(103, 23);
            this.signError.TabIndex = 21;
            this.signError.Text = "Test sign error";
            this.signError.UseVisualStyleBackColor = true;
            this.signError.Visible = false;
            this.signError.Click += new System.EventHandler(this.signError_Click);
            // 
            // signOmissionTextBox
            // 
            this.signOmissionTextBox.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.signOmissionTextBox.Font = new System.Drawing.Font("Lucida Console", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.signOmissionTextBox.ForeColor = System.Drawing.Color.Black;
            this.signOmissionTextBox.Location = new System.Drawing.Point(131, 60);
            this.signOmissionTextBox.Name = "signOmissionTextBox";
            this.signOmissionTextBox.Size = new System.Drawing.Size(23, 18);
            this.signOmissionTextBox.TabIndex = 22;
            this.signOmissionTextBox.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // digitRepeatTextBox
            // 
            this.digitRepeatTextBox.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.digitRepeatTextBox.Font = new System.Drawing.Font("Lucida Console", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.digitRepeatTextBox.ForeColor = System.Drawing.Color.Black;
            this.digitRepeatTextBox.Location = new System.Drawing.Point(131, 89);
            this.digitRepeatTextBox.Name = "digitRepeatTextBox";
            this.digitRepeatTextBox.Size = new System.Drawing.Size(23, 18);
            this.digitRepeatTextBox.TabIndex = 23;
            this.digitRepeatTextBox.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // decimalOmissionTextBox
            // 
            this.decimalOmissionTextBox.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.decimalOmissionTextBox.Font = new System.Drawing.Font("Lucida Console", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.decimalOmissionTextBox.ForeColor = System.Drawing.Color.Black;
            this.decimalOmissionTextBox.Location = new System.Drawing.Point(131, 118);
            this.decimalOmissionTextBox.Name = "decimalOmissionTextBox";
            this.decimalOmissionTextBox.Size = new System.Drawing.Size(23, 18);
            this.decimalOmissionTextBox.TabIndex = 24;
            this.decimalOmissionTextBox.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // extraDigitTextBox
            // 
            this.extraDigitTextBox.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.extraDigitTextBox.Font = new System.Drawing.Font("Lucida Console", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.extraDigitTextBox.ForeColor = System.Drawing.Color.Black;
            this.extraDigitTextBox.Location = new System.Drawing.Point(131, 147);
            this.extraDigitTextBox.Name = "extraDigitTextBox";
            this.extraDigitTextBox.Size = new System.Drawing.Size(23, 18);
            this.extraDigitTextBox.TabIndex = 25;
            this.extraDigitTextBox.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // signErrorTextBox
            // 
            this.signErrorTextBox.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.signErrorTextBox.Font = new System.Drawing.Font("Lucida Console", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.signErrorTextBox.ForeColor = System.Drawing.Color.Black;
            this.signErrorTextBox.Location = new System.Drawing.Point(131, 176);
            this.signErrorTextBox.Name = "signErrorTextBox";
            this.signErrorTextBox.Size = new System.Drawing.Size(23, 18);
            this.signErrorTextBox.TabIndex = 26;
            this.signErrorTextBox.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // misplacedDecimalTextBox
            // 
            this.misplacedDecimalTextBox.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.misplacedDecimalTextBox.Font = new System.Drawing.Font("Lucida Console", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.misplacedDecimalTextBox.ForeColor = System.Drawing.Color.Black;
            this.misplacedDecimalTextBox.Location = new System.Drawing.Point(322, 60);
            this.misplacedDecimalTextBox.Name = "misplacedDecimalTextBox";
            this.misplacedDecimalTextBox.Size = new System.Drawing.Size(23, 18);
            this.misplacedDecimalTextBox.TabIndex = 27;
            this.misplacedDecimalTextBox.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // digitOmissionTextBox
            // 
            this.digitOmissionTextBox.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.digitOmissionTextBox.Font = new System.Drawing.Font("Lucida Console", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.digitOmissionTextBox.ForeColor = System.Drawing.Color.Black;
            this.digitOmissionTextBox.Location = new System.Drawing.Point(322, 89);
            this.digitOmissionTextBox.Name = "digitOmissionTextBox";
            this.digitOmissionTextBox.Size = new System.Drawing.Size(23, 18);
            this.digitOmissionTextBox.TabIndex = 28;
            this.digitOmissionTextBox.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // wrongDigitTextBox
            // 
            this.wrongDigitTextBox.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.wrongDigitTextBox.Font = new System.Drawing.Font("Lucida Console", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.wrongDigitTextBox.ForeColor = System.Drawing.Color.Black;
            this.wrongDigitTextBox.Location = new System.Drawing.Point(322, 118);
            this.wrongDigitTextBox.Name = "wrongDigitTextBox";
            this.wrongDigitTextBox.Size = new System.Drawing.Size(23, 18);
            this.wrongDigitTextBox.TabIndex = 29;
            this.wrongDigitTextBox.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // digitTranspositionTextBox
            // 
            this.digitTranspositionTextBox.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.digitTranspositionTextBox.Font = new System.Drawing.Font("Lucida Console", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.digitTranspositionTextBox.ForeColor = System.Drawing.Color.Black;
            this.digitTranspositionTextBox.Location = new System.Drawing.Point(322, 147);
            this.digitTranspositionTextBox.Name = "digitTranspositionTextBox";
            this.digitTranspositionTextBox.Size = new System.Drawing.Size(23, 18);
            this.digitTranspositionTextBox.TabIndex = 30;
            this.digitTranspositionTextBox.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Lucida Console", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.ForeColor = System.Drawing.Color.LightGray;
            this.label3.Location = new System.Drawing.Point(3, 63);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(103, 11);
            this.label3.TabIndex = 31;
            this.label3.Text = "Sign Omission:";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Lucida Console", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label4.ForeColor = System.Drawing.Color.LightGray;
            this.label4.Location = new System.Drawing.Point(3, 92);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(96, 11);
            this.label4.TabIndex = 32;
            this.label4.Text = "Digit Repeat:";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("Lucida Console", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label5.ForeColor = System.Drawing.Color.LightGray;
            this.label5.Location = new System.Drawing.Point(3, 121);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(124, 11);
            this.label5.TabIndex = 33;
            this.label5.Text = "Decimal Omission:";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Font = new System.Drawing.Font("Lucida Console", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label6.ForeColor = System.Drawing.Color.LightGray;
            this.label6.Location = new System.Drawing.Point(3, 150);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(89, 11);
            this.label6.TabIndex = 34;
            this.label6.Text = "Extra Digit:";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Font = new System.Drawing.Font("Lucida Console", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label7.ForeColor = System.Drawing.Color.LightGray;
            this.label7.Location = new System.Drawing.Point(3, 179);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(82, 11);
            this.label7.TabIndex = 35;
            this.label7.Text = "Sign Error:";
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Font = new System.Drawing.Font("Lucida Console", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label8.ForeColor = System.Drawing.Color.LightGray;
            this.label8.Location = new System.Drawing.Point(174, 63);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(131, 11);
            this.label8.TabIndex = 36;
            this.label8.Text = "Misplaced Decimal:";
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Font = new System.Drawing.Font("Lucida Console", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label9.ForeColor = System.Drawing.Color.LightGray;
            this.label9.Location = new System.Drawing.Point(174, 92);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(110, 11);
            this.label9.TabIndex = 37;
            this.label9.Text = "Digit Omission:";
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Font = new System.Drawing.Font("Lucida Console", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label10.ForeColor = System.Drawing.Color.LightGray;
            this.label10.Location = new System.Drawing.Point(174, 121);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(89, 11);
            this.label10.TabIndex = 38;
            this.label10.Text = "Wrong Digit:";
            // 
            // label11
            // 
            this.label11.AutoSize = true;
            this.label11.Font = new System.Drawing.Font("Lucida Console", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label11.ForeColor = System.Drawing.Color.LightGray;
            this.label11.Location = new System.Drawing.Point(174, 150);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(145, 11);
            this.label11.TabIndex = 39;
            this.label11.Text = "Digit Transposition:";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(64)))), ((int)(((byte)(64)))), ((int)(((byte)(64)))));
            this.ClientSize = new System.Drawing.Size(359, 211);
            this.Controls.Add(this.label11);
            this.Controls.Add(this.label10);
            this.Controls.Add(this.label9);
            this.Controls.Add(this.label8);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.digitTranspositionTextBox);
            this.Controls.Add(this.wrongDigitTextBox);
            this.Controls.Add(this.digitOmissionTextBox);
            this.Controls.Add(this.misplacedDecimalTextBox);
            this.Controls.Add(this.signErrorTextBox);
            this.Controls.Add(this.extraDigitTextBox);
            this.Controls.Add(this.decimalOmissionTextBox);
            this.Controls.Add(this.digitRepeatTextBox);
            this.Controls.Add(this.signOmissionTextBox);
            this.Controls.Add(this.signError);
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
            this.ForeColor = System.Drawing.Color.Black;
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
        private System.Windows.Forms.Button signError;
        private System.Windows.Forms.TextBox signOmissionTextBox;
        private System.Windows.Forms.TextBox digitRepeatTextBox;
        private System.Windows.Forms.TextBox decimalOmissionTextBox;
        private System.Windows.Forms.TextBox extraDigitTextBox;
        private System.Windows.Forms.TextBox signErrorTextBox;
        private System.Windows.Forms.TextBox misplacedDecimalTextBox;
        private System.Windows.Forms.TextBox digitOmissionTextBox;
        private System.Windows.Forms.TextBox wrongDigitTextBox;
        private System.Windows.Forms.TextBox digitTranspositionTextBox;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.Label label11;
    }
}