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
        Dictionary<Excel.Workbook,List<RibbonHelper.CellColor>> color_dict; // list for storing colors
        Excel.Application app;
        Excel.Workbook current_workbook;
        bool tool_active = false;
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
            tool_active = true;
        }

        private void DeactivateTool()
        {
            this.TestNewProcedure.Enabled = true;
            this.MarkAsOK.Enabled = false;
            this.FixError.Enabled = false;
            this.clearColoringButton.Enabled = false;
            tool_active = false;
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

        // Action for "Analyze Worksheet" button
        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            current_workbook = app.ActiveWorkbook;
            try
            {
                if (current_workbook != null)
                {
                    RibbonHelper.RestoreColors2(color_dict[current_workbook], tool_highlights);
                    color_dict.Remove(current_workbook);
                }
            }
            catch { }

            if (!color_dict.ContainsKey(current_workbook))
            {
                color_dict.Add(current_workbook, RibbonHelper.SaveColors2(current_workbook));
            }

            //Disable screen updating during perturbation and analysis to speed things up
            app.ScreenUpdating = false;

            // Make a new analysisData object
            AnalysisData data = new AnalysisData(app, app.ActiveWorkbook, false);
            data.worksheets = app.Worksheets;
            data.global_stopwatch.Reset();
            data.global_stopwatch.Start();

            // Construct a new tree every time the tool is run
            data.Reset();

            //TODO This needs to be improved -- it doesn't make sense that it restores the colors to whatever they were when the workbook was opened. It removes any coloring changes the user has made since opening the file, which is not good.
            // reset colors
            //if (current_workbook != null)
            //{
            //    RibbonHelper.RestoreColors2(color_dict[current_workbook]);
            //}
            
            // Build dependency graph (modifies data)
            ConstructTree.constructTree(data, app.ActiveWorkbook, app);

            // Perturb data (modifies data)
            Analysis.perturbationAnalysis(data);
            
            // Find outliers (modifies data)
            Analysis.outlierAnalysis(data);

            string fileName = Globals.ThisAddIn.Application.ActiveWorkbook.Name;
            string folderPath = Globals.ThisAddIn.Application.ActiveWorkbook.Path;
            string reportsText = "";
            //If there is an existing report file, get its text, otherwise a new report file will be created.
            try
            {
                reportsText = System.IO.File.ReadAllText(@folderPath + @"\" + @fileName.Remove(fileName.LastIndexOf(".")) + " - Report.txt");
            }
            catch { }

            reportsText += "Worksheet Index\tAddress\tOriginal Color" + Environment.NewLine;
            foreach (string[] dataEntry in data.reportData)
            {
                foreach (string dataItem in dataEntry)
                {
                    reportsText += dataItem + "\t";
                }
                reportsText += Environment.NewLine;
            }
            if (folderPath == "")
            {
                System.Windows.Forms.MessageBox.Show("A report cannot be created because this file has not been saved yet.");
            }
            else
            {
                System.IO.File.WriteAllText(@folderPath + @"\" + @fileName.Remove(fileName.LastIndexOf(".")) + " - Report.txt", reportsText);
            }

            // Enable screen updating when we're done
            app.ScreenUpdating = true;
            if (showGVTree.Checked)
            {
                Display d = new Display();
                d.textBox1.Text = ConstructTree.GenerateGraphVizTree(data.formula_nodes);
                d.ShowDialog();
            }
        }

        // Button for outputting MTurk HIT CSVs
        private void button7_Click(object sender, RibbonControlEventArgs e)
        {
            // the longest input field we can represent on MTurk
            const int MAXLEN = 20;

            // get MTurk jobs or fail is spreadsheet data cells are too long
            TurkJob[] turkjobs;
            var turkjobs_opt = ConstructTree.DataForMTurk(current_workbook, Globals.ThisAddIn.Application, MAXLEN);
            if (FSharpOption<TurkJob[]>.get_IsSome(turkjobs_opt))
            {
                turkjobs = turkjobs_opt.Value;
            }
            else
            {
                System.Windows.Forms.MessageBox.Show("This spreadsheet contains data items with lengths longer than " + MAXLEN + " characters and cannot be converted into an MTurk job.");
                return;
            }

            // get workbook name
            var wbname = app.ActiveWorkbook.Name;

            // prompt for folder name
            var sFD = new System.Windows.Forms.FolderBrowserDialog();
            sFD.ShowDialog();

            // If the path is not an empty string, go ahead
            if (sFD.SelectedPath != "")
            {
                // write key file
                var outfile = Path.Combine(sFD.SelectedPath, wbname + ".arr");
                TurkJob.SerializeArray(outfile, turkjobs);

                // write images, 2 for each TurkJob
                foreach (TurkJob tj in turkjobs)
                {
                    tj.WriteAsImages(sFD.SelectedPath, wbname);
                }

                // write CSV
                var csvfile = Path.Combine(sFD.SelectedPath, wbname + ".csv");
                var lines = new List<string>();
                lines.Add(turkjobs[0].ToCSVHeaderLine(wbname));
                lines.AddRange(turkjobs.Select(turkjob => turkjob.ToCSVLine(wbname)));
                File.WriteAllLines(csvfile, lines);
            }
        }

        private TreeScore Analyze()
        {
            current_workbook = app.ActiveWorkbook;

            // Disable screen updating during perturbation and analysis to speed things up
            app.ScreenUpdating = false;

            // Make a new analysisData object
            AnalysisData data = new AnalysisData(app, app.ActiveWorkbook, false);
            data.worksheets = app.Worksheets;

            // Construct a new tree every time the tool is run
            data.Reset();

            // Build dependency graph (modifies data)
            ConstructTree.constructTree(data, app.ActiveWorkbook, app);

            if (data.TerminalInputNodes().Length == 0)
            {
                System.Windows.Forms.MessageBox.Show("There are no ranges that can be analyzed in this spreadsheet.");
                data.pb.Close();
                app.ScreenUpdating = true;
                return new TreeScore();
            }

            // e * 1000
            var NBOOTS = (int)(Math.Ceiling(1000 * Math.Exp(1.0)));

            // Get bootstraps
            var scores = Analysis.Bootstrap(NBOOTS, data, app, this.weighted.Checked);

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

        private void performanceExperiments_Click(object sender, RibbonControlEventArgs e)
        {
            PerformanceExperiments experimentsForm = new PerformanceExperiments();
            experimentsForm.ShowDialog();
        }

        private void countFormulas_Click(object sender, RibbonControlEventArgs e)
        {
            int countFormulas = 0;
            foreach (Excel.Range cell in current_workbook.ActiveSheet.UsedRange)
            {
                if (cell.HasFormula)
                {
                    countFormulas++;
                }
            }
            System.Windows.Forms.MessageBox.Show(countFormulas + " formulas in this workbook.");
        }
        private void undoButton_Click(object sender, RibbonControlEventArgs e)
        {
            string fileName = Globals.ThisAddIn.Application.ActiveWorkbook.Name;
            string folderPath = Globals.ThisAddIn.Application.ActiveWorkbook.Path;
            string reportsText = "";
            try
            {
                reportsText = System.IO.File.ReadAllText(@folderPath + @"\" + @fileName.Remove(fileName.LastIndexOf(".")) + " - Report.txt");
            }
            catch 
            { 
                return; 
            }

            int startIndex = reportsText.LastIndexOf("Worksheet Index\tAddress\tOriginal Color" + Environment.NewLine);
            
            //If the reports file is empty, there is nothing more to undo
            if (startIndex == -1)
            {
                return;
            }
            //Restore colors will go here

            string lastReport = reportsText.Substring(startIndex);
            //string[] lastReportLines = lastReport.lin
            System.Windows.Forms.MessageBox.Show("Last report: " + Environment.NewLine + lastReport);
            reportsText = reportsText.Remove(startIndex);
            System.Windows.Forms.MessageBox.Show("Remaining report: "  + reportsText);
            System.IO.File.WriteAllText(@folderPath + @"\" + @fileName.Remove(fileName.LastIndexOf(".")) + " - Report.txt", reportsText);
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
    }
}
