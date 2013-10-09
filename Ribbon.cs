using System;
using System.Collections.Generic;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.Office.Tools.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using DataDebugMethods;
using TreeNode = DataDebugMethods.TreeNode;
using TreeScore = System.Collections.Generic.Dictionary<DataDebugMethods.TreeNode, int>;
using ColorDict = System.Collections.Generic.Dictionary<Microsoft.Office.Interop.Excel.Workbook, System.Collections.Generic.List<DataDebugMethods.TreeNode>>;
using TreeDict = System.Collections.Generic.Dictionary<AST.Address, DataDebugMethods.TreeNode>;
using Microsoft.FSharp.Core;
using System.IO;
using System.Linq;

namespace DataDebug
{
    public partial class Ribbon
    {
        // e * 1000
        public readonly static int NBOOTS = (int)(Math.Ceiling(1000 * Math.Exp(1.0)));

        Dictionary<Excel.Workbook,List<RibbonHelper.CellColor>> color_dict; // list for storing colors
        Excel.Application app;
        Excel.Workbook current_workbook;
        HashSet<AST.Address> tool_highlights = new HashSet<AST.Address>();
        HashSet<AST.Address> known_good = new HashSet<AST.Address>();
        TreeScore analysis_results = null;
        AST.Address flagged_cell = null;
        System.Drawing.Color GREEN = System.Drawing.Color.Green;

        private void ActivateTool()
        {
            this.MarkAsOK.Enabled = true;
            this.FixError.Enabled = true;
            this.clearColoringButton.Enabled = true;
            this.TestNewProcedure.Enabled = false;
        }

        private void DeactivateTool()
        {
            this.TestNewProcedure.Enabled = true;
            this.MarkAsOK.Enabled = false;
            this.FixError.Enabled = false;
            this.clearColoringButton.Enabled = false;
        }

        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            // start tool in deactivated state
            DeactivateTool();

            // init color storage
            color_dict = new Dictionary<Excel.Workbook, List<RibbonHelper.CellColor>>();

            // Get current app
            app = Globals.ThisAddIn.Application;

            // Get current workbook
            current_workbook = app.ActiveWorkbook;

            // save colors
            if (current_workbook != null)
            {
                color_dict.Add(current_workbook, RibbonHelper.SaveColors2(current_workbook));
            }

            // register event handlers
            app.WorkbookOpen += new Microsoft.Office.Interop.Excel.AppEvents_WorkbookOpenEventHandler(app_WorkbookOpen);
            app.WorkbookBeforeClose += new Microsoft.Office.Interop.Excel.AppEvents_WorkbookBeforeCloseEventHandler(app_WorkbookBeforeClose);
            app.WorkbookActivate += new Microsoft.Office.Interop.Excel.AppEvents_WorkbookActivateEventHandler(app_WorkbookActivate);
        }

        private void app_WorkbookOpen(Excel.Workbook wb)
        {
            current_workbook = wb;
            if (!color_dict.ContainsKey(current_workbook))
            {
                color_dict.Add(current_workbook, RibbonHelper.SaveColors2(current_workbook));
            }
        }

        void app_WorkbookBeforeClose(Excel.Workbook wb, ref bool cancel)
        {
            color_dict.Remove(wb);
            if (current_workbook == wb)
            {
                current_workbook = null;
            }
        }

        void app_WorkbookActivate(Excel.Workbook wb)
        {
            current_workbook = wb;
            if (!color_dict.ContainsKey(current_workbook))
            {
                color_dict.Add(current_workbook, RibbonHelper.SaveColors2(current_workbook));
            }
        }

        private TreeScore Analyze()
        {
            current_workbook = app.ActiveWorkbook;

            // Disable screen updating during analysis to speed things up
            app.ScreenUpdating = false;

            // Make a new analysisData object
            AnalysisData data = new AnalysisData(app, app.ActiveWorkbook, false);

            // Build dependency graph (modifies data)
            ConstructTree.constructTree(data, app.ActiveWorkbook, app);

            if (data.TerminalInputNodes().Length == 0)
            {
                System.Windows.Forms.MessageBox.Show("This spreadsheet contains no functions that take inputs.");
                data.KillPB();
                app.ScreenUpdating = true;
                return new TreeScore();
            }

            // Get bootstraps
            var scores = Analysis.Bootstrap(NBOOTS, data, app, true);

            // Color top outlier and save in ribbon state
            flagged_cell = Analysis.FlagTopOutlier(scores, known_good);
            if (flagged_cell == null)
            {
                System.Windows.Forms.MessageBox.Show("No bugs remain.");
                ResetTool();
            }
            else
            {
                tool_highlights.Add(flagged_cell);

                // enable auditing buttons
                ActivateTool();
            }

            // Enable screen updating when we're done
            app.ScreenUpdating = true;

            return scores;
        }

        private void TestNewProcedure_Click(object sender, RibbonControlEventArgs e)
        {
            analysis_results = Analyze();
        }

        private void ResetTool()
        {
            if (current_workbook != null)
            {
                RibbonHelper.RestoreColors2(color_dict[current_workbook], tool_highlights);
            }

            known_good.Clear();
            tool_highlights.Clear();
            DeactivateTool();
        }

        // Action for "Clear coloring" button
        private void clearColoringButton_Click(object sender, RibbonControlEventArgs e)
        {
            ResetTool();
        }

        private void MarkAsOK_Click(object sender, RibbonControlEventArgs e)
        {
            known_good.Add(flagged_cell);
            var cell = flagged_cell.GetCOMObject(app);
            cell.Interior.Color = GREEN;
            flagged_cell = Analysis.FlagTopOutlier(analysis_results, known_good);
            if (flagged_cell == null)
            {
                System.Windows.Forms.MessageBox.Show("No bugs remain.");
                ResetTool();
            }
        }

        private void FixError_Click(object sender, RibbonControlEventArgs e)
        {
            var comcell = flagged_cell.GetCOMObject(app);
            System.Action callback = () => {
                flagged_cell = null;
                analysis_results = Analyze();
            };
            var fixform = new CellFixForm(comcell, GREEN, callback);
            fixform.Show();
        }

        private void TestStuff_Click(object sender, RibbonControlEventArgs e)
        {
            double[] a = { 1, 2, 2, 2, 3, 1, 5, 6, 6, 6, 0 };
            var result = DataDebugMethods.Analysis.ComputeQuantile(a);
            System.Windows.Forms.MessageBox.Show(String.Join(",", result));
        }
    }
}
