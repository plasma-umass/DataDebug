namespace DataDebug
{
    partial class Ribbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

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

        #region Component Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl1 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl2 = this.Factory.CreateRibbonDropDownItem();
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.button1 = this.Factory.CreateRibbonButton();
            this.checkBox2 = this.Factory.CreateRibbonCheckBox();
            this.dropDown1 = this.Factory.CreateRibbonDropDown();
            this.button7 = this.Factory.CreateRibbonButton();
            this.button8 = this.Factory.CreateRibbonButton();
            this.button2 = this.Factory.CreateRibbonButton();
            this.toggle_compile_regex = this.Factory.CreateRibbonCheckBox();
            this.toggle_global_perturbation = this.Factory.CreateRibbonCheckBox();
            this.toggle_analyze_outliers = this.Factory.CreateRibbonCheckBox();
            this.peirce_button = this.Factory.CreateRibbonButton();
            this.toggle_weighted_average = this.Factory.CreateRibbonCheckBox();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.button3 = this.Factory.CreateRibbonButton();
            this.button4 = this.Factory.CreateRibbonButton();
            this.button5 = this.Factory.CreateRibbonButton();
            this.button6 = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.group2.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.group1);
            this.tab1.Groups.Add(this.group2);
            this.tab1.Label = "TabAddIns";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.button1);
            this.group1.Items.Add(this.checkBox2);
            this.group1.Items.Add(this.dropDown1);
            this.group1.Items.Add(this.button7);
            this.group1.Items.Add(this.button8);
            this.group1.Items.Add(this.button2);
            this.group1.Items.Add(this.toggle_compile_regex);
            this.group1.Items.Add(this.toggle_global_perturbation);
            this.group1.Items.Add(this.toggle_analyze_outliers);
            this.group1.Items.Add(this.peirce_button);
            this.group1.Items.Add(this.toggle_weighted_average);
            this.group1.Label = "DataDebug";
            this.group1.Name = "group1";
            // 
            // button1
            // 
            this.button1.Label = "Analyze Document";
            this.button1.Name = "button1";
            this.button1.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button1_Click);
            // 
            // checkBox2
            // 
            this.checkBox2.Label = "Fuzz Repeated Values Simultaneously";
            this.checkBox2.Name = "checkBox2";
            this.checkBox2.Visible = false;
            // 
            // dropDown1
            // 
            ribbonDropDownItemImpl1.Label = "Column";
            ribbonDropDownItemImpl2.Label = "Row";
            this.dropDown1.Items.Add(ribbonDropDownItemImpl1);
            this.dropDown1.Items.Add(ribbonDropDownItemImpl2);
            this.dropDown1.Label = "Data Layout";
            this.dropDown1.Name = "dropDown1";
            this.dropDown1.Visible = false;
            // 
            // button7
            // 
            this.button7.Label = "button7";
            this.button7.Name = "button7";
            this.button7.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button7_Click);
            // 
            // button8
            // 
            this.button8.Label = "Clear Coloring";
            this.button8.Name = "button8";
            this.button8.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button8_Click);
            // 
            // button2
            // 
            this.button2.Label = "Derivatives";
            this.button2.Name = "button2";
            this.button2.Visible = false;
            this.button2.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button2_Click);
            // 
            // toggle_compile_regex
            // 
            this.toggle_compile_regex.Checked = true;
            this.toggle_compile_regex.Label = "Compile Regular Expressions";
            this.toggle_compile_regex.Name = "toggle_compile_regex";
            this.toggle_compile_regex.Visible = false;
            // 
            // toggle_global_perturbation
            // 
            this.toggle_global_perturbation.Label = "Global Perturbation";
            this.toggle_global_perturbation.Name = "toggle_global_perturbation";
            this.toggle_global_perturbation.Visible = false;
            // 
            // toggle_analyze_outliers
            // 
            this.toggle_analyze_outliers.Label = "Highlight Important Outliers Only";
            this.toggle_analyze_outliers.Name = "toggle_analyze_outliers";
            this.toggle_analyze_outliers.Visible = false;
            // 
            // peirce_button
            // 
            this.peirce_button.Label = "Peirce Criterion for selected range";
            this.peirce_button.Name = "peirce_button";
            this.peirce_button.Visible = false;
            this.peirce_button.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.peirce_button_Click);
            // 
            // toggle_weighted_average
            // 
            this.toggle_weighted_average.Label = "Look for outliers in weighted average z-score";
            this.toggle_weighted_average.Name = "toggle_weighted_average";
            this.toggle_weighted_average.Visible = false;
            // 
            // group2
            // 
            this.group2.Items.Add(this.button3);
            this.group2.Items.Add(this.button4);
            this.group2.Items.Add(this.button5);
            this.group2.Items.Add(this.button6);
            this.group2.Label = "Statistics";
            this.group2.Name = "group2";
            this.group2.Visible = false;
            // 
            // button3
            // 
            this.button3.Label = "Normal Anderson-Darling Test";
            this.button3.Name = "button3";
            this.button3.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button3_Click);
            // 
            // button4
            // 
            this.button4.Label = "Find Outliers";
            this.button4.Name = "button4";
            this.button4.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button4_Click);
            // 
            // button5
            // 
            this.button5.Label = "Clear";
            this.button5.Name = "button5";
            this.button5.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button5_Click);
            // 
            // button6
            // 
            this.button6.Label = "Normal KS Test";
            this.button6.Name = "button6";
            this.button6.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button6_Click);
            // 
            // Ribbon
            // 
            this.Name = "Ribbon";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown dropDown1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button2;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button3;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button4;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button5;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button6;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox checkBox2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button7;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button8;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox toggle_compile_regex;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox toggle_global_perturbation;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox toggle_analyze_outliers;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton peirce_button;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox toggle_weighted_average;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon Ribbon1
        {
            get { return this.GetRibbon<Ribbon>(); }
        }
    }
}