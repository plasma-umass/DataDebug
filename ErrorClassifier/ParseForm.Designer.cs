namespace ErrorClassifier
{
    partial class ParseForm
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
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.button1 = new System.Windows.Forms.Button();
            this.chompButton = new System.Windows.Forms.Button();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.folderBrowserDialog1 = new System.Windows.Forms.FolderBrowserDialog();
            this.selectFolder = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.generateFuzzed = new System.Windows.Forms.Button();
            this.textBox2 = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.runTool = new System.Windows.Forms.Button();
            this.label6 = new System.Windows.Forms.Label();
            this.textBox3 = new System.Windows.Forms.TextBox();
            this.label7 = new System.Windows.Forms.Label();
            this.numericUpDown1 = new System.Windows.Forms.NumericUpDown();
            this.strawMan = new System.Windows.Forms.Button();
            this.strawManCheckBox = new System.Windows.Forms.CheckBox();
            this.oldAnalysis = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDown1)).BeginInit();
            this.SuspendLayout();
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.DefaultExt = "csv";
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(1202, 34);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(125, 23);
            this.button1.TabIndex = 0;
            this.button1.Text = "Browse for csv file";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Visible = false;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // chompButton
            // 
            this.chompButton.Location = new System.Drawing.Point(1252, 5);
            this.chompButton.Name = "chompButton";
            this.chompButton.Size = new System.Drawing.Size(75, 23);
            this.chompButton.TabIndex = 1;
            this.chompButton.Text = "Chomp";
            this.chompButton.UseVisualStyleBackColor = true;
            this.chompButton.Visible = false;
            this.chompButton.Click += new System.EventHandler(this.chompButton_Click);
            // 
            // textBox1
            // 
            this.textBox1.AcceptsReturn = true;
            this.textBox1.AcceptsTab = true;
            this.textBox1.BackColor = System.Drawing.Color.Black;
            this.textBox1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.textBox1.Font = new System.Drawing.Font("Lucida Console", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.textBox1.ForeColor = System.Drawing.Color.LightGray;
            this.textBox1.Location = new System.Drawing.Point(12, 66);
            this.textBox1.Multiline = true;
            this.textBox1.Name = "textBox1";
            this.textBox1.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.textBox1.Size = new System.Drawing.Size(871, 452);
            this.textBox1.TabIndex = 2;
            this.textBox1.WordWrap = false;
            this.textBox1.TextChanged += new System.EventHandler(this.textBox1_TextChanged);
            // 
            // selectFolder
            // 
            this.selectFolder.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(64)))), ((int)(((byte)(64)))), ((int)(((byte)(64)))));
            this.selectFolder.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.selectFolder.Font = new System.Drawing.Font("Lucida Console", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.selectFolder.ForeColor = System.Drawing.Color.LightGray;
            this.selectFolder.Location = new System.Drawing.Point(74, 37);
            this.selectFolder.Name = "selectFolder";
            this.selectFolder.Size = new System.Drawing.Size(175, 23);
            this.selectFolder.TabIndex = 3;
            this.selectFolder.Text = "Select Folder";
            this.selectFolder.UseVisualStyleBackColor = false;
            this.selectFolder.Click += new System.EventHandler(this.selectFolder_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Lucida Console", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.ForeColor = System.Drawing.Color.LightGray;
            this.label1.Location = new System.Drawing.Point(12, 12);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(358, 16);
            this.label1.TabIndex = 4;
            this.label1.Text = "Generate Fuzzed Files and Run Tool:";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Lucida Console", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.ForeColor = System.Drawing.Color.LightGray;
            this.label2.Location = new System.Drawing.Point(14, 41);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(54, 11);
            this.label2.TabIndex = 5;
            this.label2.Text = "Step 1:";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Lucida Console", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.ForeColor = System.Drawing.Color.LightGray;
            this.label3.Location = new System.Drawing.Point(267, 41);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(54, 11);
            this.label3.TabIndex = 6;
            this.label3.Text = "Step 2:";
            // 
            // generateFuzzed
            // 
            this.generateFuzzed.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.generateFuzzed.Font = new System.Drawing.Font("Lucida Console", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.generateFuzzed.ForeColor = System.Drawing.Color.LightGray;
            this.generateFuzzed.Location = new System.Drawing.Point(327, 37);
            this.generateFuzzed.Name = "generateFuzzed";
            this.generateFuzzed.Size = new System.Drawing.Size(175, 23);
            this.generateFuzzed.TabIndex = 7;
            this.generateFuzzed.Text = "Generate Fuzz Files";
            this.generateFuzzed.UseVisualStyleBackColor = true;
            this.generateFuzzed.Click += new System.EventHandler(this.generateFuzzed_Click);
            // 
            // textBox2
            // 
            this.textBox2.AcceptsReturn = true;
            this.textBox2.AcceptsTab = true;
            this.textBox2.BackColor = System.Drawing.Color.Black;
            this.textBox2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.textBox2.Font = new System.Drawing.Font("Lucida Console", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.textBox2.ForeColor = System.Drawing.Color.LightGray;
            this.textBox2.Location = new System.Drawing.Point(889, 66);
            this.textBox2.Multiline = true;
            this.textBox2.Name = "textBox2";
            this.textBox2.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.textBox2.Size = new System.Drawing.Size(438, 192);
            this.textBox2.TabIndex = 8;
            this.textBox2.WordWrap = false;
            this.textBox2.TextChanged += new System.EventHandler(this.textBox2_TextChanged);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Lucida Console", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label4.ForeColor = System.Drawing.Color.LightGray;
            this.label4.Location = new System.Drawing.Point(886, 47);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(278, 16);
            this.label4.TabIndex = 9;
            this.label4.Text = "Error Classification Table:";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("Lucida Console", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label5.ForeColor = System.Drawing.Color.LightGray;
            this.label5.Location = new System.Drawing.Point(528, 41);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(54, 11);
            this.label5.TabIndex = 10;
            this.label5.Text = "Step 3:";
            // 
            // runTool
            // 
            this.runTool.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.runTool.Font = new System.Drawing.Font("Lucida Console", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.runTool.ForeColor = System.Drawing.Color.LightGray;
            this.runTool.Location = new System.Drawing.Point(588, 37);
            this.runTool.Name = "runTool";
            this.runTool.Size = new System.Drawing.Size(157, 23);
            this.runTool.TabIndex = 11;
            this.runTool.Text = "Run Bootstrap";
            this.runTool.UseVisualStyleBackColor = true;
            this.runTool.Click += new System.EventHandler(this.runTool_Click);
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Font = new System.Drawing.Font("Lucida Console", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label6.ForeColor = System.Drawing.Color.LightGray;
            this.label6.Location = new System.Drawing.Point(886, 270);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(188, 16);
            this.label6.TabIndex = 12;
            this.label6.Text = "Detection Results:";
            // 
            // textBox3
            // 
            this.textBox3.BackColor = System.Drawing.Color.Black;
            this.textBox3.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.textBox3.ForeColor = System.Drawing.Color.LightGray;
            this.textBox3.Location = new System.Drawing.Point(889, 289);
            this.textBox3.Multiline = true;
            this.textBox3.Name = "textBox3";
            this.textBox3.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.textBox3.Size = new System.Drawing.Size(438, 229);
            this.textBox3.TabIndex = 13;
            this.textBox3.TextChanged += new System.EventHandler(this.textBox3_TextChanged);
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.ForeColor = System.Drawing.Color.LightGray;
            this.label7.Location = new System.Drawing.Point(585, 15);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(79, 13);
            this.label7.TabIndex = 14;
            this.label7.Text = "NBOOTS = E x";
            // 
            // numericUpDown1
            // 
            this.numericUpDown1.BackColor = System.Drawing.Color.Black;
            this.numericUpDown1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.numericUpDown1.ForeColor = System.Drawing.Color.LightGray;
            this.numericUpDown1.Location = new System.Drawing.Point(670, 12);
            this.numericUpDown1.Maximum = new decimal(new int[] {
            10000,
            0,
            0,
            0});
            this.numericUpDown1.Minimum = new decimal(new int[] {
            50,
            0,
            0,
            0});
            this.numericUpDown1.Name = "numericUpDown1";
            this.numericUpDown1.Size = new System.Drawing.Size(75, 20);
            this.numericUpDown1.TabIndex = 16;
            this.numericUpDown1.Value = new decimal(new int[] {
            1000,
            0,
            0,
            0});
            // 
            // strawMan
            // 
            this.strawMan.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.strawMan.Font = new System.Drawing.Font("Lucida Console", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.strawMan.ForeColor = System.Drawing.Color.LightGray;
            this.strawMan.Location = new System.Drawing.Point(889, 11);
            this.strawMan.Name = "strawMan";
            this.strawMan.Size = new System.Drawing.Size(117, 23);
            this.strawMan.TabIndex = 17;
            this.strawMan.Text = "Straw Man";
            this.strawMan.UseVisualStyleBackColor = true;
            this.strawMan.Click += new System.EventHandler(this.strawMan_Click);
            // 
            // strawManCheckBox
            // 
            this.strawManCheckBox.AutoSize = true;
            this.strawManCheckBox.Checked = true;
            this.strawManCheckBox.CheckState = System.Windows.Forms.CheckState.Checked;
            this.strawManCheckBox.Font = new System.Drawing.Font("Lucida Console", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.strawManCheckBox.ForeColor = System.Drawing.Color.LightGray;
            this.strawManCheckBox.Location = new System.Drawing.Point(751, 40);
            this.strawManCheckBox.Name = "strawManCheckBox";
            this.strawManCheckBox.Size = new System.Drawing.Size(129, 15);
            this.strawManCheckBox.TabIndex = 18;
            this.strawManCheckBox.Text = "+Straw Man Test";
            this.strawManCheckBox.UseVisualStyleBackColor = true;
            // 
            // oldAnalysis
            // 
            this.oldAnalysis.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.oldAnalysis.Font = new System.Drawing.Font("Lucida Console", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.oldAnalysis.ForeColor = System.Drawing.Color.LightGray;
            this.oldAnalysis.Location = new System.Drawing.Point(1012, 11);
            this.oldAnalysis.Name = "oldAnalysis";
            this.oldAnalysis.Size = new System.Drawing.Size(117, 23);
            this.oldAnalysis.TabIndex = 19;
            this.oldAnalysis.Text = "Run Old Tool";
            this.oldAnalysis.UseVisualStyleBackColor = true;
            this.oldAnalysis.Click += new System.EventHandler(this.oldAnalysis_Click);
            // 
            // ParseForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(64)))), ((int)(((byte)(64)))), ((int)(((byte)(64)))));
            this.ClientSize = new System.Drawing.Size(1332, 530);
            this.Controls.Add(this.oldAnalysis);
            this.Controls.Add(this.strawManCheckBox);
            this.Controls.Add(this.strawMan);
            this.Controls.Add(this.numericUpDown1);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.textBox3);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.runTool);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.textBox2);
            this.Controls.Add(this.generateFuzzed);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.selectFolder);
            this.Controls.Add(this.textBox1);
            this.Controls.Add(this.chompButton);
            this.Controls.Add(this.button1);
            this.Name = "ParseForm";
            this.Text = "ParseForm";
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDown1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Button chompButton;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog1;
        private System.Windows.Forms.Button selectFolder;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Button generateFuzzed;
        private System.Windows.Forms.TextBox textBox2;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Button runTool;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.TextBox textBox3;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.NumericUpDown numericUpDown1;
        private System.Windows.Forms.Button strawMan;
        private System.Windows.Forms.CheckBox strawManCheckBox;
        private System.Windows.Forms.Button oldAnalysis;
    }
}