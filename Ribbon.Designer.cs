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
            this.tab1 = this.Factory.CreateRibbonTab();
            this.ccgroup = this.Factory.CreateRibbonGroup();
            this.TestNewProcedure = this.Factory.CreateRibbonButton();
            this.MarkAsOK = this.Factory.CreateRibbonButton();
            this.FixError = this.Factory.CreateRibbonButton();
            this.clearColoringButton = this.Factory.CreateRibbonButton();
            this.TestStuff = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.ccgroup.SuspendLayout();
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
            this.ccgroup.Items.Add(this.TestNewProcedure);
            this.ccgroup.Items.Add(this.MarkAsOK);
            this.ccgroup.Items.Add(this.FixError);
            this.ccgroup.Items.Add(this.clearColoringButton);
            this.ccgroup.Items.Add(this.TestStuff);
            this.ccgroup.Label = "CheckCell";
            this.ccgroup.Name = "ccgroup";
            // 
            // TestNewProcedure
            // 
            this.TestNewProcedure.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.TestNewProcedure.Image = global::DataDebug.Properties.Resources.analyze_small;
            this.TestNewProcedure.Label = "Analyze";
            this.TestNewProcedure.Name = "TestNewProcedure";
            this.TestNewProcedure.ShowImage = true;
            this.TestNewProcedure.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.TestNewProcedure_Click);
            // 
            // MarkAsOK
            // 
            this.MarkAsOK.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.MarkAsOK.Image = global::DataDebug.Properties.Resources.mark_as_ok_small;
            this.MarkAsOK.Label = "Mark As OK";
            this.MarkAsOK.Name = "MarkAsOK";
            this.MarkAsOK.ShowImage = true;
            this.MarkAsOK.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.MarkAsOK_Click);
            // 
            // FixError
            // 
            this.FixError.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.FixError.Image = global::DataDebug.Properties.Resources.correct_small;
            this.FixError.Label = "Fix Error";
            this.FixError.Name = "FixError";
            this.FixError.ShowImage = true;
            this.FixError.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.FixError_Click);
            // 
            // clearColoringButton
            // 
            this.clearColoringButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.clearColoringButton.Image = global::DataDebug.Properties.Resources.clear_small;
            this.clearColoringButton.Label = "Start Over";
            this.clearColoringButton.Name = "clearColoringButton";
            this.clearColoringButton.ShowImage = true;
            this.clearColoringButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.clearColoringButton_Click);
            // 
            // TestStuff
            // 
            this.TestStuff.Label = "Test";
            this.TestStuff.Name = "TestStuff";
            this.TestStuff.Visible = false;
            this.TestStuff.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.TestStuff_Click);
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

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton clearColoringButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton TestNewProcedure;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup ccgroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton MarkAsOK;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton FixError;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton TestStuff;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon Ribbon1
        {
            get { return this.GetRibbon<Ribbon>(); }
        }
    }
}