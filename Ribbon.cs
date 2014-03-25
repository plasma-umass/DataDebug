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
        public readonly static long MAX_DURATION_IN_MS = 5L * 60L * 1000L;  // 5 minutes

        Dictionary<Excel.Workbook,List<RibbonHelper.CellColor>> color_dict; // list for storing colors
        Excel.Application app;
        Excel.Workbook current_workbook;
        double tool_significance = 0.95;
        HashSet<AST.Address> tool_highlights = new HashSet<AST.Address>();
        HashSet<AST.Address> output_highlights = new HashSet<AST.Address>();
        HashSet<AST.Address> known_good = new HashSet<AST.Address>();
        //IEnumerable<Tuple<double, TreeNode>> analysis_results = null;
        List<KeyValuePair<TreeNode, int>> flaggable_list = null;
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
            //    //Normal per range shown; others hidden
            //    ccgroup.Visible = false;
            //    group1.Visible = true;
            //    group2.Visible = false;
            //}
            //else
            //{
            //    //Normal on all inputs shown; others hidden
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

        private List<KeyValuePair<TreeNode, int>> Analyze(bool normal_cutoff, long max_duration_in_ms)
        {
            var sw = new System.Diagnostics.Stopwatch();
            sw.Start();

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
                    //return new List<Tuple<double, TreeNode>>();
                    return new List<KeyValuePair<TreeNode, int>>();
                }

                // Get bootstraps
                var scores = Analysis.Bootstrap(NBOOTS, data, app, true, true, max_duration_in_ms, sw);
                var scores_list = scores.OrderByDescending(pair => pair.Value).ToList();

                List<KeyValuePair<TreeNode, int>> filtered_high_scores = null;

                //Using an outlier test for highlighting 
                //scores that fall outside of two standard deviations from the others
                //The one-sided 5% cutoff for the normal distribution is 1.6448.

                if (normal_cutoff)
                {
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

                    if (tool_significance == 0.95)
                    {
                        filtered_high_scores = scores_list.Where(kvp => kvp.Value > mean + std_deviation * 1.6448).ToList();
                    }
                    else if (tool_significance == 0.9)   //10% cutoff 1.2815
                    {
                        filtered_high_scores = scores_list.Where(kvp => kvp.Value > mean + std_deviation * 1.2815).ToList();
                    }
                    else if (tool_significance == 0.975) //2.5% cutoff 1.9599
                    {
                        filtered_high_scores = scores_list.Where(kvp => kvp.Value > mean + std_deviation * 1.9599).ToList();
                    }
                    else if (tool_significance == 0.925) //7.5% cutoff 1.4395
                    {
                        filtered_high_scores = scores_list.Where(kvp => kvp.Value > mean + std_deviation * 1.4395).ToList();
                    }
                }
                else
                {
                    int start_ptr = 0;
                    int end_ptr = 0;

                    List<KeyValuePair<TreeNode, int>> high_scores = new List<KeyValuePair<TreeNode, int>>();

                    while ((double)start_ptr / scores_list.Count < 1.0 - tool_significance) //the start of this score region is before the cutoff
                    {
                        //while the scores at the start and end pointers are the same, bump the end pointer
                        while (end_ptr < scores_list.Count && scores_list[start_ptr].Value == scores_list[end_ptr].Value)
                        {
                            end_ptr++;
                        }
                        //Now the end_pointer points to the first index with a lower score
                        //If the number of entries with the current value is fewer than the significance cutoff, add all values of this score to the high_scores list; the number of entries is equal to the end_ptr since end_ptr is zero-based
                        //There is some added "wiggle room" to the cutoff, so that the last entry is allowed to straddle the cutoff bound.
                        //  To do this, we add (1 / total number of entries) to the cutoff
                        //The purpose of the wiggle room is to allow us to deal with small ranges (less than 20 entries), since a single entry accounts
                        //for more than 5% of the total.
                        //      Note: tool_significance is along the lines of 0.95 (not 0.05).
                        if ((double)end_ptr / scores_list.Count < 1.0 - tool_significance + (double)1.0 / scores_list.Count)
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
                    filtered_high_scores = high_scores.Where(kvp => !known_good.Contains(kvp.Key.GetAddress())).ToList();
                }
                // Enable screen updating when we're done
                app.ScreenUpdating = true;

                sw.Stop();

                return filtered_high_scores;
            }
        }

        private void ActivateAndCenterOn(AST.Address cell, Excel.Application app)
        {
            // go to worksheet
            RibbonHelper.GetWorksheetByName(cell.A1Worksheet(), current_workbook.Worksheets).Activate();

            // COM object
            var comobj = cell.GetCOMObject(app);

            // center screen on cell
            var visible_columns = app.ActiveWindow.VisibleRange.Columns.Count;
            var visible_rows = app.ActiveWindow.VisibleRange.Rows.Count;
            app.Goto(comobj, true);
            app.ActiveWindow.SmallScroll(Type.Missing, visible_rows / 2, Type.Missing, visible_columns / 2);

            // select highlighted cell
            // center on highlighted cell
            comobj.Select();

        }

        private void Flag()
        {
            //filter known_good
            flaggable_list = flaggable_list.Where(kvp => !known_good.Contains(kvp.Key.GetAddress())).ToList();
            if (flaggable_list.Count() != 0)
            {
                // get TreeNode corresponding to most unusual score
                flagged_cell = flaggable_list[0].Key.GetAddress();
            }
            else
            {
                flagged_cell = null;
            }

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

                flagged_cell.GetCOMObject(app).Interior.Color = System.Drawing.Color.Red;
                tool_highlights.Add(flagged_cell);

                // go to highlighted cell
                ActivateAndCenterOn(flagged_cell, app);

                // enable auditing buttons
                ActivateTool();
            }
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
                    flaggable_list = Analyze(false, MAX_DURATION_IN_MS);
                    Flag();
                }
                catch (ExcelParserUtility.ParseException ex)
                {
                    System.Windows.Forms.Clipboard.SetText(ex.Message);
                    System.Windows.Forms.MessageBox.Show("Could not parse the formula string:\n" + ex.Message);
                    return;
                }
            }
        }

        //This clears the outputs highlighted in yellow
        private void RestoreOutputColors()
        {
            if (current_workbook != null)
            {
                RibbonHelper.RestoreColors2(color_dict[current_workbook], output_highlights);
            }
            output_highlights.Clear();
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
            RestoreOutputColors();
            Flag();
        }

        private void FixError_Click(object sender, RibbonControlEventArgs e)
        {
            var comcell = flagged_cell.GetCOMObject(app);
            System.Action callback = () => {
                flagged_cell = null;
                try
                {
                    flaggable_list = Analyze(false, MAX_DURATION_IN_MS);
                    Flag();
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
            RestoreOutputColors();
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
                output_highlights.Add(node.GetAddress());
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
                    switch (analysisType.SelectedItemIndex)
                    {
                        case 0:
                            sim.Run(2700, filename, 0.95, app, 0.05, c, rng, UserSimulation.AnalysisType.CheckCell, true, false, false, MAX_DURATION_IN_MS);
                            break;
                        case 1:
                            sim.Run(2700, filename, 0.95, app, 0.05, c, rng, UserSimulation.AnalysisType.NormalPerRange, true, false, false, MAX_DURATION_IN_MS);
                            break;
                        case 2:
                            sim.Run(2700, filename, 0.95, app, 0.05, c, rng, UserSimulation.AnalysisType.NormalAllInputs, true, false, false, MAX_DURATION_IN_MS);
                            break;
                        default:
                            System.Windows.Forms.MessageBox.Show("There was a problem with the selection of analysis type.");
                            break;
                    }
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

        private void ToDOT_Click(object sender, RibbonControlEventArgs e)
        {
            var data = ConstructTree.constructTree(app.ActiveWorkbook, app);
            var graph = data.ToDOT();
            System.Windows.Forms.Clipboard.SetText(graph);
            System.Windows.Forms.MessageBox.Show("In clipboard");
        }

        private void LoopCheck_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                var data = ConstructTree.constructTree(app.ActiveWorkbook, app);
                var contains_loop = data.ContainsLoop();
                System.Windows.Forms.MessageBox.Show("Contains loops: " + contains_loop);
            }
            catch (ExcelParserUtility.ParseException ex)
            {
                System.Windows.Forms.Clipboard.SetText(ex.Message);
                System.Windows.Forms.MessageBox.Show(String.Format("Parser exception for formula: {0}", ex.Message));
            }
        }
    }
}
