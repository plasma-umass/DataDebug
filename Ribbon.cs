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
        double tool_significance = 0.95;
        HashSet<AST.Address> tool_highlights = new HashSet<AST.Address>();
        HashSet<AST.Address> known_good = new HashSet<AST.Address>();
        IEnumerable<Tuple<double, TreeNode>> analysis_results = null;
        AST.Address flagged_cell = null;
        System.Drawing.Color GREEN = System.Drawing.Color.Green;
        string classification_file;

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

        private FSharpOption<double> GetSignificance(string input, string label)
        {
            var errormsg = label + " must be a value between 0 and 100";
            var significance = 0.95;

            try
            {
                significance = (100.0 - Double.Parse(input)) / 100.0;
            }
            catch
            {
                System.Windows.Forms.MessageBox.Show(errormsg);
            }

            if (significance < 0 || significance > 100)
            {
                System.Windows.Forms.MessageBox.Show(errormsg);
            }

            return FSharpOption<double>.Some(significance);
        }

        private IEnumerable<Tuple<double,TreeNode>> Analyze()
        {
            current_workbook = app.ActiveWorkbook;

            // Disable screen updating during analysis to speed things up
            app.ScreenUpdating = false;

            // Build dependency graph (modifies data)
            var data = ConstructTree.constructTree(app.ActiveWorkbook, app, true);

            if (data.TerminalInputNodes().Length == 0)
            {
                System.Windows.Forms.MessageBox.Show("This spreadsheet contains no functions that take inputs.");
                data.KillPB();
                app.ScreenUpdating = true;
                return new List<Tuple<double, TreeNode>>();
            }

            // Get bootstraps
            var scores = Analysis.Bootstrap(NBOOTS, data, app, true);

            // Compute quantiles based on user-supplied sensitivity
            var quantiles = Analysis.ComputeQuantile<int, TreeNode>(scores.Select(
                pair => new Tuple<int, TreeNode>(pair.Value, pair.Key))
            );

            // Color top outlier, zoom to worksheet, and save in ribbon state
            flagged_cell = Analysis.FlagTopOutlier(quantiles, known_good, tool_significance, app);
            if (flagged_cell == null)
            {
                System.Windows.Forms.MessageBox.Show("No bugs remain.");
                ResetTool();
            }
            else
            {
                tool_highlights.Add(flagged_cell);

                // go to highlighted cell
                ActivateAndCenterOn(flagged_cell, app);

                // enable auditing buttons
                ActivateTool();
            }

            // Enable screen updating when we're done
            app.ScreenUpdating = true;

            return quantiles;
        }

        private void ActivateAndCenterOn(AST.Address cell, Excel.Application app)
        {
            // go to worksheet
            RibbonHelper.GetWorksheetByName(flagged_cell.A1Worksheet(), current_workbook.Worksheets).Activate();

            // COM object
            var comobj = flagged_cell.GetCOMObject(app);

            // center screen on cell
            var visible_columns = app.ActiveWindow.VisibleRange.Columns.Count;
            var visible_rows = app.ActiveWindow.VisibleRange.Rows.Count;
            app.Goto(comobj, true);
            app.ActiveWindow.SmallScroll(Type.Missing, visible_rows / 2, Type.Missing, visible_columns / 2);

            // select highlighted cell
            // center on highlighted cell
            comobj.Select();

        }

        private void TestNewProcedure_Click(object sender, RibbonControlEventArgs e)
        {
            var sig = GetSignificance(this.SensitivityTextBox.Text, this.SensitivityTextBox.Label);
            if (sig == FSharpOption<double>.None)
            {
                return;
            }
            else
            {
                tool_significance = sig.Value;
                analysis_results = Analyze();
            }
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
            flagged_cell = Analysis.FlagTopOutlier(analysis_results, known_good, tool_significance, app);
            if (flagged_cell == null)
            {
                System.Windows.Forms.MessageBox.Show("No bugs remain.");
                ResetTool();
            }
            else
            {
                // go to highlighted cell
                ActivateAndCenterOn(flagged_cell, app);
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
            RunSimulation_Click(sender, e);
        }

        private void RunSimulation_Click(object sender, RibbonControlEventArgs e)
        {
            // the full path of this workbook
            var filename = app.ActiveWorkbook.FullName;

            // the default output filename
            var r = new System.Text.RegularExpressions.Regex(@"(.+)\.xls|xlsx", System.Text.RegularExpressions.RegexOptions.Compiled);
            var default_output_file = r.Match(app.ActiveWorkbook.Name).Groups[1].Value + "_sim_results.csv";

            // init a RNG
            var rng = new Random();

            // ask the user for the classification data
            var ofd = new System.Windows.Forms.OpenFileDialog();
            ofd.FileName = "ClassificationData-2013-11-14.bin";
            if (ofd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                // deserialize classification
                var c = UserSimulation.Classification.Deserialize(ofd.FileName);

                // ask the user where the output data should go
                var sfd = new System.Windows.Forms.SaveFileDialog();
                sfd.FileName = default_output_file;

                if (sfd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    // run the simulation
                    app.ActiveWorkbook.Close(false, Type.Missing, Type.Missing);    // why?
                    UserSimulation.Simulation sim = new UserSimulation.Simulation();
                    sim.Run(2700, filename, 0.95, app, 0.05, c, rng);
                    sim.ToCSV(sfd.FileName);
                }
            }
        }

        private void ErrorBtn_Click(object sender, RibbonControlEventArgs e)
        {
            // open classifier file
            if (classification_file == null)
            {
                var ofd = new System.Windows.Forms.OpenFileDialog();
                if (ofd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    classification_file = ofd.FileName;
                }
            }
            var c = UserSimulation.Classification.Deserialize(classification_file);
            var egen = new UserSimulation.ErrorGenerator();

            // get cursor
            var cursor = app.ActiveCell;

            // get string at current cursor
            String data = cursor.Value2;

            // get error string
            String baddata = egen.GenerateErrorString(data, c);

            // put string back into spreadsheet
            cursor.Value2 = baddata;
        }
    }
}
