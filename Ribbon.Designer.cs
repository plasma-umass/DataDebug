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
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl3 = this.Factory.CreateRibbonDropDownItem();
            this.tab1 = this.Factory.CreateRibbonTab();
            this.ccgroup = this.Factory.CreateRibbonGroup();
            this.buttonGroup1 = this.Factory.CreateRibbonButtonGroup();
            this.TestNewProcedure = this.Factory.CreateRibbonButton();
            this.MarkAsOK = this.Factory.CreateRibbonButton();
            this.FixError = this.Factory.CreateRibbonButton();
            this.clearColoringButton = this.Factory.CreateRibbonButton();
            this.SensitivityTextBox = this.Factory.CreateRibbonEditBox();
            this.RunSimulation = this.Factory.CreateRibbonButton();
            this.TestStuff = this.Factory.CreateRibbonButton();
            this.ErrorBtn = this.Factory.CreateRibbonButton();
            this.analysisType = this.Factory.CreateRibbonDropDown();
            this.tab1.SuspendLayout();
            this.ccgroup.SuspendLayout();
            this.buttonGroup1.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.ccgroup);
            this.tab1.Label = "TabAddIns";
            this.tab1.Name = "tab1";
            // 
            // ccgroup
            // 
            this.ccgroup.Items.Add(this.buttonGroup1);
            this.ccgroup.Items.Add(this.SensitivityTextBox);
            this.ccgroup.Items.Add(this.RunSimulation);
            this.ccgroup.Items.Add(this.ErrorBtn);
            this.ccgroup.Items.Add(this.TestStuff);
            this.ccgroup.Items.Add(this.analysisType);
            this.ccgroup.Label = "CheckCell";
            this.ccgroup.Name = "ccgroup";
            // 
            // buttonGroup1
            // 
            this.buttonGroup1.Items.Add(this.TestNewProcedure);
            this.buttonGroup1.Items.Add(this.MarkAsOK);
            this.buttonGroup1.Items.Add(this.FixError);
            this.buttonGroup1.Items.Add(this.clearColoringButton);
            this.buttonGroup1.Name = "buttonGroup1";
            // 
            // TestNewProcedure
            // 
            this.TestNewProcedure.Image = global::DataDebug.Properties.Resources.analyze_small;
            this.TestNewProcedure.Label = "Analyze";
            this.TestNewProcedure.Name = "TestNewProcedure";
            this.TestNewProcedure.ShowImage = true;
            this.TestNewProcedure.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.TestNewProcedure_Click);
            // 
            // MarkAsOK
            // 
            this.MarkAsOK.Image = global::DataDebug.Properties.Resources.mark_as_ok_small;
            this.MarkAsOK.Label = "Mark as OK";
            this.MarkAsOK.Name = "MarkAsOK";
            this.MarkAsOK.ShowImage = true;
            this.MarkAsOK.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.MarkAsOK_Click);
            // 
            // FixError
            // 
            this.FixError.Image = global::DataDebug.Properties.Resources.correct_small;
            this.FixError.Label = "Fix Error";
            this.FixError.Name = "FixError";
            this.FixError.ShowImage = true;
            this.FixError.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.FixError_Click);
            // 
            // clearColoringButton
            // 
            this.clearColoringButton.Image = global::DataDebug.Properties.Resources.clear_small;
            this.clearColoringButton.Label = "Start Over";
            this.clearColoringButton.Name = "clearColoringButton";
            this.clearColoringButton.ShowImage = true;
            this.clearColoringButton.Visible = false;
            this.clearColoringButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.clearColoringButton_Click);
            // 
            // SensitivityTextBox
            // 
            this.SensitivityTextBox.Label = "% Most Unusual to Show";
            this.SensitivityTextBox.Name = "SensitivityTextBox";
            this.SensitivityTextBox.Text = "5.0";
            // 
            // RunSimulation
            // 
            this.RunSimulation.Label = "Run Simulation";
            this.RunSimulation.Name = "RunSimulation";
            this.RunSimulation.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.RunSimulation_Click);
            // 
            // TestStuff
            // 
            this.TestStuff.Label = "Test";
            this.TestStuff.Name = "TestStuff";
            this.TestStuff.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.TestStuff_Click);
            // 
            // ErrorBtn
            // 
            this.ErrorBtn.Label = "Make Error";
            this.ErrorBtn.Name = "ErrorBtn";
            this.ErrorBtn.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ErrorBtn_Click);
            // analysisType
            // 
            ribbonDropDownItemImpl1.Label = "CheckCell";
            ribbonDropDownItemImpl2.Label = "Normal (per range)";
            ribbonDropDownItemImpl3.Label = "Normal (all inputs)";
            this.analysisType.Items.Add(ribbonDropDownItemImpl1);
            this.analysisType.Items.Add(ribbonDropDownItemImpl2);
            this.analysisType.Items.Add(ribbonDropDownItemImpl3);
            this.analysisType.Label = "Analysis Type";
            this.analysisType.Name = "analysisType";
            // 
            // Ribbon
            // 
            this.Name = "Ribbon";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.ccgroup.ResumeLayout(false);
            this.ccgroup.PerformLayout();
            this.buttonGroup1.ResumeLayout(false);
            this.buttonGroup1.PerformLayout();

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton clearColoringButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton TestNewProcedure;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup ccgroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton MarkAsOK;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton FixError;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton TestStuff;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox SensitivityTextBox;
        internal Microsoft.Office.Tools.Ribbon.RibbonButtonGroup buttonGroup1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton RunSimulation;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ErrorBtn;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown analysisType;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon Ribbon1
        {
            get { return this.GetRibbon<Ribbon>(); }
        }
    }
}