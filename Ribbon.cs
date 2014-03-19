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
        AnalysisData data;

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
            ////Randomly select the version of the tool that should be shown;
            //int tool_versions = 3;
            //Random rand = new Random();
            //int i = rand.Next(tool_versions);
            //i = 0;

            //if (i == 0)
            //{
            //    //CheckCell shown; others hidden
            //    ccgroup.Visible = true;
            //    group1.Visible = false;
            //    group2.Visible = false;
            //}
            //else if (i == 1)
            //{
            //    //Normal shown; others hidden
            //    ccgroup.Visible = false;
            //    group1.Visible = true;
            //    group2.Visible = false;
            //}
            //else
            //{
            //    //Grubb's shown; others hidden
            //    ccgroup.Visible = false;
            //    group1.Visible = false;
            //    group2.Visible = true;
            //}

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
            using (var pb = new ProgBar(0, 100))
            {
                current_workbook = app.ActiveWorkbook;

                // Disable screen updating during analysis to speed things up
                app.ScreenUpdating = false;

                // Build dependency graph (modifies data)
                try
                {
                    data = ConstructTree.constructTree(app.ActiveWorkbook, app, pb);
                }
                catch (ExcelParserUtility.ParseException e)
                {
                    // cleanup UI and then rethrow
                    app.ScreenUpdating = true;
                    throw e;
                }

                if (data.TerminalInputNodes().Length == 0)
                {
                    System.Windows.Forms.MessageBox.Show("This spreadsheet contains no functions that take inputs.");
                    app.ScreenUpdating = true;
                    return new List<Tuple<double, TreeNode>>();
                }

                // Get bootstraps
                var scores = Analysis.Bootstrap(NBOOTS, data, app, true, false);
                var scores_list = scores.OrderByDescending(pair => pair.Value).ToList(); //pair => pair.Key, pair => pair.Value);

                //Should we be using an outlier test for 
                //highlighting scores that fall outside of two standard deviations from the others?
                //(The one-sided 5% cutoff for the normal distribution is 1.6448.)
                //Or do we always want to be highlighting the top 5% of the scores?
                //Currently, if we have something like this, we don't flag anything 
                //because the 1 value that is weird is an entire 20% of the total:
                //     1,1,1,1000,1     =SUM(A1:A5)
                //-Dimitar

                // TODO: don't forget that we never want to flag a cell that failed
                // zero hypothesis tests.

                //Code for doing normal outlier analysis on the scores:
                //find mean:
                double sum = 0.0;
                foreach (double d in scores.Values)
                {
                    sum += d;
                }
                double mean = sum / scores.Values.Count;
                //find variance
                double distance_sum_sq = 0.0;
                foreach (double d in scores.Values)
                {
                    distance_sum_sq += Math.Pow(mean - d, 2);
                }
                double variance = distance_sum_sq / scores.Values.Count;

                //find std. deviation
                double std_deviation = Math.Sqrt(variance);

                //scores_list = scores_list.Where

                int start_ptr = 0;
                int end_ptr = 0;

                var weird_scores = scores_list.Where(kvp => kvp.Value > mean + std_deviation * 1.6448).ToList();
                List<KeyValuePair<TreeNode, int>> high_scores = new List<KeyValuePair<TreeNode, int>>();

                while ((double)start_ptr / scores_list.Count < 1.0 - tool_significance) //the start of this score region is before the cutoff
                {
                    //while the scores at the start and end pointers are the same, bump the end pointer
                    while (end_ptr < scores_list.Count && scores_list[start_ptr].Value == scores_list[end_ptr].Value)
                    {
                        end_ptr++;
                    }
                    //Now the end_pointer points to the first index with a lower score
                    //If the end pointer is still above the significance cutoff, add all values of this score to the high_scores list
                    if ((double)end_ptr / scores_list.Count < 1.0 - tool_significance)
                    {
                        //add all values of the current score to high_scores list
                        for (; start_ptr < end_ptr; start_ptr++)
                        {
                            high_scores.Add(scores_list[start_ptr]);
                        }
                        //Increment the start pointer to the start of the next score region
                        start_ptr++;
                    }
                    else    //if this score region extends past the cutoff, we don't add any of its values to the high_scores list, and stop
                    {
                        break;
                    }
                }

                // filter out cells marked as OK
                var filtered_scores = high_scores.Where(kvp => !known_good.Contains(kvp.Key.GetAddress())).ToList();
                AST.Address flagged_cell;
                if (filtered_scores.Count() != 0)
                {
                    // get TreeNode corresponding to most unusual score
                    flagged_cell = filtered_scores[0].Key.GetAddress();
                }
                else
                {
                    flagged_cell = null;
                }

                // Compute quantiles based on user-supplied sensitivity
                //            var quantiles = Analysis.ComputeQuantile<int, TreeNode>(scores.Select(
                //                pair => new Tuple<int, TreeNode>(pair.Value, pair.Key))
                //            );

                // Color top outlier, zoom to worksheet, and save in ribbon state
                //TODO color in yellow the outputs that depend on the outlier being flagged
                //            flagged_cell = Analysis.FlagTopOutlier(quantiles, known_good, tool_significance, app);
                if (flagged_cell == null)
                {
                    System.Windows.Forms.MessageBox.Show("No bugs remain.");
                    ResetTool();
                }
                else
                {
                    TreeNode flagged_node;
                    data.cell_nodes.TryGetValue(flagged_cell, out flagged_node);

                    if (flagged_node.hasOutputs())
                    {
                        foreach (var output in flagged_node.getOutputs())
                        {
                            exploreNode(output);
                        }
                    }

                    tool_highlights.Add(flagged_cell);

                    // go to highlighted cell
                    ActivateAndCenterOn(flagged_cell, app);

                    // enable auditing buttons
                    ActivateTool();
                }

                // Enable screen updating when we're done
                app.ScreenUpdating = true;
                return null;
                //return quantiles;
            }
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
                try
                {
                    analysis_results = Analyze();
                }
                catch (ExcelParserUtility.ParseException ex)
                {
                    System.Windows.Forms.Clipboard.SetText(ex.Message);
                    System.Windows.Forms.MessageBox.Show("Could not parse the formula string:\n" + ex.Message);
                    return;
                }
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
                TreeNode flagged_node;
                data.cell_nodes.TryGetValue(flagged_cell, out flagged_node);

                if (flagged_node.hasOutputs())
                {
                    foreach (var output in flagged_node.getOutputs())
                    {
                        exploreNode(output);
                    }
                }

                tool_highlights.Add(flagged_cell);

                // go to highlighted cell
                ActivateAndCenterOn(flagged_cell, app);
            }
        }

        private void FixError_Click(object sender, RibbonControlEventArgs e)
        {
            var comcell = flagged_cell.GetCOMObject(app);
            System.Action callback = () => {
                flagged_cell = null;
                try
                {
                    analysis_results = Analyze();
                }
                catch (ExcelParserUtility.ParseException ex)
                {
                    System.Windows.Forms.Clipboard.SetText(ex.Message);
                    System.Windows.Forms.MessageBox.Show("Could not parse the formula string:\n" + ex.Message);
                    return;
                }
            };
            var fixform = new CellFixForm(comcell, GREEN, callback);
            fixform.Show();
        }

        //Recursive method for highlighting the outputs reachable from a certain TreeNode
        private void exploreNode(TreeNode node)
        {
            if (node.hasOutputs())
            {
                foreach (var o in node.getOutputs())
                {
                    exploreNode(o);
                }
            }
            else
            {
                node.getCOMObject().Interior.Color = System.Drawing.Color.Yellow;
            }
        }

        private void TestStuff_Click(object sender, RibbonControlEventArgs e)
        {
            using (var pb = new ProgBar(0,100)) {
                System.Windows.Forms.MessageBox.Show("" + analysisType.SelectedItem);

                //RunSimulation_Click(sender, e);
                var sig = GetSignificance(this.SensitivityTextBox.Text, this.SensitivityTextBox.Label);
                if (sig == FSharpOption<double>.None)
                {
                    return;
                }
                else
                {
                    tool_significance = sig.Value;
                }

                current_workbook = app.ActiveWorkbook;

                // Disable screen updating during analysis to speed things up
                app.ScreenUpdating = false;

                // Build dependency graph (modifies data)
                AnalysisData data;
                try
                {
                    data = ConstructTree.constructTree(app.ActiveWorkbook, app, pb);
                }
                catch (ExcelParserUtility.ParseException ex)
                {
                    // cleanup UI and rethrow
                    app.ScreenUpdating = true;
                    throw ex;
                }

                if (data.TerminalInputNodes().Length == 0)
                {
                    System.Windows.Forms.MessageBox.Show("This spreadsheet contains no functions that take inputs.");
                    app.ScreenUpdating = true;
                    return;
                }

                //foreach (var node in data.TerminalFormulaNodes())
                //{
                //    node.getCOMObject().Interior.Color = System.Drawing.Color.Yellow;
                //}
                var tin = data.TerminalInputNodes();

                foreach (var input_range in data.TerminalInputNodes())
                {
                    foreach (var input_node in input_range.getInputs())
                    {
                        //find the ultimate outputs for this input
                        if (input_node.hasOutputs())
                        {
                            foreach (var output in input_node.getOutputs())
                            {
                                exploreNode(output);
                            }
                        }
                    }
                }

                foreach (var range in data.input_ranges.Values)
                {
                    var normal_dist = new DataDebugMethods.NormalDistribution(range.getCOMObject());

                    for (int error_index = 0; error_index < normal_dist.errorsCount(); error_index++)
                    {
                        normal_dist.getError(error_index).Interior.Color = System.Drawing.Color.Violet;
                    }
                }
                /**
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
                **/
                // Enable screen updating when we're done
                app.ScreenUpdating = true;
            }
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
                    sim.Run(2700, filename, 0.95, app, 0.05, c, rng, analysisType.SelectedItem.ToString());
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


            if (classification_file != null)
            {
                var c = UserSimulation.Classification.Deserialize(classification_file);
                var egen = new UserSimulation.ErrorGenerator();

                // get cursor
                var cursor = app.ActiveCell;

                // get string at current cursor
                String data = System.Convert.ToString(cursor.Value2);

                // get error string
                String baddata = egen.GenerateErrorString(data, c);

                // put string back into spreadsheet
                cursor.Value2 = baddata;
            }
        }
    }
}
