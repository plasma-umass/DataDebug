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
using OptTuple = Microsoft.FSharp.Core.FSharpOption<System.Tuple<UserSimulation.Classification, string>>;

namespace DataDebug
{
    public partial class Ribbon
    {
        #region OLDCODE
        //Dictionary<Excel.Workbook,List<RibbonHelper.CellColor>> color_dict; // list for storing colors
        //Excel.Application app;
        //Excel.Workbook current_workbook;
        //double tool_significance = 0.95;
        //HashSet<AST.Address> tool_highlights = new HashSet<AST.Address>();
        //HashSet<AST.Address> output_highlights = new HashSet<AST.Address>();
        //HashSet<AST.Address> known_good = new HashSet<AST.Address>();
        //IEnumerable<Tuple<double, TreeNode>> analysis_results = null;
        //List<KeyValuePair<TreeNode, int>> flaggable_list = null;
        //AST.Address flagged_cell = null;
        //AnalysisData data;

        ////Recursive method for highlighting the outputs reachable from a certain TreeNode
        //private void exploreNode(TreeNode node)
        //{
        //    if (node.hasOutputs())
        //    {
        //        foreach (var o in node.getOutputs())
        //        {
        //            exploreNode(o);
        //        }
        //    }
        //    else
        //    {
        //        node.getCOMObject().Interior.Color = System.Drawing.Color.Yellow;
        //        output_highlights.Add(node.GetAddress());
        //    }
        //}
        #endregion OLDCODE

        // workbook state data
        Dictionary<Excel.Workbook, WorkbookState> wbstates = new Dictionary<Excel.Workbook, WorkbookState>();
        WorkbookState current_workbook;

        // simulation files
        string classification_file;
        String benchmark_dir;
        String simulation_output_dir;
        String simulation_classification_file;

        private void SetUIState(WorkbookState wbs) {
            this.MarkAsOK.Enabled = wbs.MarkAsOK_Enabled;
            this.FixError.Enabled = wbs.FixError_Enabled;
            this.clearColoringButton.Enabled = wbs.ClearColoringButton_Enabled;
            this.Analyze.Enabled = wbs.Analyze_Enabled;
        }

        
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            // Callbacks for handling workbook state objects
            WorkbookOpen(Globals.ThisAddIn.Application.ActiveWorkbook);
            ((Excel.AppEvents_Event)Globals.ThisAddIn.Application).NewWorkbook += WorkbookOpen;
            Globals.ThisAddIn.Application.WorkbookOpen += WorkbookOpen;
            Globals.ThisAddIn.Application.WorkbookActivate += WorkbookActivated;
            Globals.ThisAddIn.Application.WorkbookDeactivate += WorkbookDeactivated;
            Globals.ThisAddIn.Application.WorkbookBeforeClose += WorkbookClose;
        }

        // This event is called when Excel opens a workbook
        private void WorkbookOpen(Excel.Workbook workbook)
        {
            wbstates.Add(workbook, new WorkbookState(Globals.ThisAddIn.Application, workbook));
        }

        // This event is called when Excel brings an opened workbook
        // to the foreground
        private void WorkbookActivated(Excel.Workbook workbook)
        {
            current_workbook = wbstates[workbook];
            SetUIState(current_workbook);
        }

        // This even it called when Excel sends an opened workbook
        // to the background
        private void WorkbookDeactivated(Excel.Workbook workbook)
        {
            current_workbook = null;
        }

        private void WorkbookClose(Excel.Workbook workbook, ref bool Cancel)
        {
            wbstates.Remove(workbook);
        }

        #region BUTTON_HANDLERS
        private void Analyze_Click(object sender, RibbonControlEventArgs e)
        {
            var sig = GetSignificance(this.SensitivityTextBox.Text, this.SensitivityTextBox.Label);
            if (sig == FSharpOption<double>.None)
            {
                return;
            }
            else
            {
                current_workbook.ToolSignificance = sig.Value;
                try
                {
                    current_workbook.Analyze(WorkbookState.MAX_DURATION_IN_MS);
                    current_workbook.Flag();
                }
                catch (ExcelParserUtility.ParseException ex)
                {
                    System.Windows.Forms.Clipboard.SetText(ex.Message);
                    System.Windows.Forms.MessageBox.Show("Could not parse the formula string:\n" + ex.Message);
                    return;
                }
            }
        }

        // Action for "Clear coloring" button
        private void clearColoringButton_Click(object sender, RibbonControlEventArgs e)
        {
            current_workbook.ResetTool();
        }

        private void MarkAsOK_Click(object sender, RibbonControlEventArgs e)
        {
            System.Windows.Forms.MessageBox.Show("foo3");
            //known_good.Add(flagged_cell);
            //var cell = flagged_cell.GetCOMObject(app);
            //cell.Interior.Color = GREEN;
            //RestoreOutputColors();
            //Flag();
        }

        private void FixError_Click(object sender, RibbonControlEventArgs e)
        {
            System.Windows.Forms.MessageBox.Show("foo4");
            //var comcell = flagged_cell.GetCOMObject(app);
            //System.Action callback = () => {
            //    flagged_cell = null;
            //    try
            //    {
            //        flaggable_list = Analyze(false, MAX_DURATION_IN_MS);
            //        Flag();
            //    }
            //    catch (ExcelParserUtility.ParseException ex)
            //    {
            //        System.Windows.Forms.Clipboard.SetText(ex.Message);
            //        System.Windows.Forms.MessageBox.Show("Could not parse the formula string:\n" + ex.Message);
            //        return;
            //    }
            //};
            //var fixform = new CellFixForm(comcell, GREEN, callback);
            //fixform.Show();
            //RestoreOutputColors();
        }

        private void TestStuff_Click(object sender, RibbonControlEventArgs e)
        {
            System.Windows.Forms.MessageBox.Show("foo5");
            //// init a RNG
            //var rng = new Random();

            //// classification data
            //UserSimulation.Classification c;

            //// ask the user for the classification data
            //if (simulation_classification_file == null)
            //{
            //    var ofd = new System.Windows.Forms.OpenFileDialog();
            //    ofd.ShowHelp = true;
            //    ofd.FileName = "ClassificationData-2013-11-14.bin";
            //    ofd.Title = "Please select a classification data input file.";
            //    if (ofd.ShowDialog() != System.Windows.Forms.DialogResult.OK)
            //    {
            //        return;
            //    }

            //    simulation_classification_file = ofd.FileName;
            //}

            //// deserialize classification
            //c = UserSimulation.Classification.Deserialize(simulation_classification_file);

            //// ask the user where to find the input data
            //if (benchmark_dir == null)
            //{
            //    var cdd = new System.Windows.Forms.FolderBrowserDialog();
            //    cdd.Description = "Please choose the folder containing the benchmark data.";
            //    if (cdd.ShowDialog() != System.Windows.Forms.DialogResult.OK)
            //    {
            //        return;
            //    }
            //    benchmark_dir = cdd.SelectedPath;
            //}

            //// enumerate files in benchmark_dir
            //var benchmark_filenames = Directory.EnumerateFiles(benchmark_dir, "*.xls", SearchOption.AllDirectories);

            //// ask the user where the output data should go
            //if (simulation_output_dir == null)
            //{
            //    var cdd = new System.Windows.Forms.FolderBrowserDialog();
            //    cdd.Description = "Please choose an output folder.";
            //    if (cdd.ShowDialog() != System.Windows.Forms.DialogResult.OK)
            //    {
            //        return;
            //    }
            //    simulation_output_dir = cdd.SelectedPath;
            //}

            //// get sig values
            //var thresh = 0.05;
            //Double.TryParse(this.SensitivityTextBox.Text, out thresh);

            //// calculate progress bar bounds
            //var pb_min = 0;
            //var pb_max = 100 * benchmark_filenames.Count();

            //// show progress bar
            //var pb = new ProgBar(pb_min, pb_max);
            //pb.Show();

            //foreach (string benchmark in benchmark_filenames)
            //{
            //    try
            //    {
            //        // open workbook
            //        Excel.Workbook wb = Utility.OpenWorkbook(benchmark, app);

            //        // run simulation
            //        RunSimulations(app, wb, rng, c, simulation_output_dir, thresh, pb);

            //        // close workbook
            //        wb.Close();
            //    }
            //    catch (Exception)
            //    {
            //        // do nothing
            //    }
            //}

            //// close progbar
            //pb.Close();

            //// inform user
            //System.Windows.Forms.MessageBox.Show(String.Format("Analysis complete.  Results are in {0}", simulation_output_dir));
        }

        private void RunSimulation_Click(object sender, RibbonControlEventArgs e)
        {
            System.Windows.Forms.MessageBox.Show("foo6");
            //// init a RNG
            //var rng = new Random();

            //// classification data
            //UserSimulation.Classification c;

            //// ask the user for the classification data
            //if (simulation_classification_file == null)
            //{
            //    var ofd = new System.Windows.Forms.OpenFileDialog();
            //    ofd.ShowHelp = true;
            //    ofd.FileName = "ClassificationData-2013-11-14.bin";
            //    ofd.Title = "Please select a classification data input file.";
            //    if (ofd.ShowDialog() != System.Windows.Forms.DialogResult.OK)
            //    {
            //        return;
            //    }

            //    simulation_classification_file = ofd.FileName;
            //}

            //// deserialize classification
            //c = UserSimulation.Classification.Deserialize(simulation_classification_file);

            //// ask the user where the output data should go
            //if (simulation_output_dir == null)
            //{
            //    var cdd = new System.Windows.Forms.FolderBrowserDialog();
            //    cdd.Description = "Please choose an output folder.";
            //    if (cdd.ShowDialog() != System.Windows.Forms.DialogResult.OK)
            //    {
            //        return;
            //    }
            //    simulation_output_dir = cdd.SelectedPath;
            //}

            //// get sig values
            //var thresh = 0.05;
            //Double.TryParse(this.SensitivityTextBox.Text, out thresh);

            //// show progress bar
            //var pb = new ProgBar(0, 100);
            //pb.Show();

            //// run simulation
            //RunSimulations(app, current_workbook, rng, c, simulation_output_dir, thresh, pb);

            //// close progbar
            //pb.Close();

            //// inform user
            //System.Windows.Forms.MessageBox.Show(String.Format("Analysis complete.  Results are in {0}", simulation_output_dir));
        }



        private void ErrorBtn_Click(object sender, RibbonControlEventArgs e)
        {
            System.Windows.Forms.MessageBox.Show("foo7");

            //// open classifier file
            //if (classification_file == null)
            //{
            //    var ofd = new System.Windows.Forms.OpenFileDialog();
            //    ofd.ShowHelp = true;
            //    if (ofd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            //    {
            //        classification_file = ofd.FileName;
            //    }
            //}

            //if (classification_file != null)
            //{
            //    var c = UserSimulation.Classification.Deserialize(classification_file);
            //    var egen = new UserSimulation.ErrorGenerator();

            //    // get cursor
            //    var cursor = app.ActiveCell;

            //    // get string at current cursor
            //    String data = System.Convert.ToString(cursor.Value2);

            //    // get error string
            //    String baddata = egen.GenerateErrorString(data, c);

            //    // put string back into spreadsheet
            //    cursor.Value2 = baddata;
            //}
        }

        private void ToDOT_Click(object sender, RibbonControlEventArgs e)
        {
            System.Windows.Forms.MessageBox.Show("foo8");

            //var data = ConstructTree.constructTree(app.ActiveWorkbook, app);
            //var graph = data.ToDOT();
            //System.Windows.Forms.Clipboard.SetText(graph);
            //System.Windows.Forms.MessageBox.Show("In clipboard");
        }

        private void LoopCheck_Click(object sender, RibbonControlEventArgs e)
        {
            System.Windows.Forms.MessageBox.Show("foo9");

            //try
            //{
            //    var data = ConstructTree.constructTree(app.ActiveWorkbook, app);
            //    var contains_loop = data.ContainsLoop();
            //    System.Windows.Forms.MessageBox.Show("Contains loops: " + contains_loop);
            //}
            //catch (ExcelParserUtility.ParseException ex)
            //{
            //    System.Windows.Forms.Clipboard.SetText(ex.Message);
            //    System.Windows.Forms.MessageBox.Show(String.Format("Parser exception for formula: {0}", ex.Message));
            //}
        }

        private void RunReviewerExperiment_Click(object sender, RibbonControlEventArgs e)
        {
            System.Windows.Forms.MessageBox.Show("foo10");

            //// init a RNG
            //var rng = new Random();

            //// classification data
            //UserSimulation.Classification c;

            //// ask the user for the classification data
            //if (simulation_classification_file == null)
            //{
            //    var ofd = new System.Windows.Forms.OpenFileDialog();
            //    ofd.ShowHelp = true;
            //    ofd.FileName = "ClassificationData-2013-11-14.bin";
            //    ofd.Title = "Please select a classification data input file.";
            //    if (ofd.ShowDialog() != System.Windows.Forms.DialogResult.OK)
            //    {
            //        return;
            //    }

            //    simulation_classification_file = ofd.FileName;
            //}

            //// deserialize classification
            //c = UserSimulation.Classification.Deserialize(simulation_classification_file);

            //// ask the user where the output data should go
            //if (simulation_output_dir == null)
            //{
            //    var cdd = new System.Windows.Forms.FolderBrowserDialog();
            //    cdd.Description = "Please choose an output folder.";
            //    if (cdd.ShowDialog() != System.Windows.Forms.DialogResult.OK)
            //    {
            //        return;
            //    }
            //    simulation_output_dir = cdd.SelectedPath;
            //}

            //// get sig values
            //var thresh = 0.05;
            //Double.TryParse(this.SensitivityTextBox.Text, out thresh);

            //// show progress bar
            //var pb = new ProgBar(0, 100);
            //pb.Show();

            //// run simulation
            //RunProportionExperiment(app, current_workbook, rng, c, simulation_output_dir, thresh, pb);

            //// close progbar
            //pb.Close();

            //// inform user
            //System.Windows.Forms.MessageBox.Show(String.Format("Analysis complete.  Results are in {0}", simulation_output_dir));
        }

        private void RunAllRevSim_Click(object sender, RibbonControlEventArgs e)
        {
            System.Windows.Forms.MessageBox.Show("foo11");

            //// init a RNG
            //var rng = new Random();

            //// classification data
            //UserSimulation.Classification c;

            //// ask the user for the classification data
            //if (simulation_classification_file == null)
            //{
            //    var ofd = new System.Windows.Forms.OpenFileDialog();
            //    ofd.ShowHelp = true;
            //    ofd.FileName = "ClassificationData-2013-11-14.bin";
            //    ofd.Title = "Please select a classification data input file.";
            //    if (ofd.ShowDialog() != System.Windows.Forms.DialogResult.OK)
            //    {
            //        return;
            //    }

            //    simulation_classification_file = ofd.FileName;
            //}

            //// deserialize classification
            //c = UserSimulation.Classification.Deserialize(simulation_classification_file);

            //// ask the user where to find the input data
            //if (benchmark_dir == null)
            //{
            //    var cdd = new System.Windows.Forms.FolderBrowserDialog();
            //    cdd.Description = "Please choose the folder containing the benchmark data.";
            //    if (cdd.ShowDialog() != System.Windows.Forms.DialogResult.OK)
            //    {
            //        return;
            //    }
            //    benchmark_dir = cdd.SelectedPath;
            //}

            //// enumerate files in benchmark_dir
            //var benchmark_filenames = Directory.EnumerateFiles(benchmark_dir, "*.xls", SearchOption.AllDirectories);

            //// ask the user where the output data should go
            //if (simulation_output_dir == null)
            //{
            //    var cdd = new System.Windows.Forms.FolderBrowserDialog();
            //    cdd.Description = "Please choose an output folder.";
            //    if (cdd.ShowDialog() != System.Windows.Forms.DialogResult.OK)
            //    {
            //        return;
            //    }
            //    simulation_output_dir = cdd.SelectedPath;
            //}

            //// get sig values
            //var thresh = 0.05;
            //Double.TryParse(this.SensitivityTextBox.Text, out thresh);

            //// calculate progress bar bounds
            //var pb_min = 0;
            //var pb_max = 100 * benchmark_filenames.Count();

            //// show progress bar
            //var pb = new ProgBar(pb_min, pb_max);
            //pb.Show();

            //// filter to pick up where we left off
            ////var files = benchmark_filenames.Where(fn => String.Compare(System.IO.Path.GetFileName(fn), "month") >= 0);
            //var files = benchmark_filenames;

            //foreach (string benchmark in files)
            //{
            //    try
            //    {
            //        // open workbook
            //        Excel.Workbook wb = Utility.OpenWorkbook(benchmark, app);

            //        // run simulation
            //        RunProportionExperiment(app, wb, rng, c, simulation_output_dir, thresh, pb);

            //        // close workbook
            //        wb.Close();
            //    }
            //    catch (Exception)
            //    {
            //        // do nothing
            //    }
            //}

            //// close progbar
            //pb.Close();

            //// inform user
            //System.Windows.Forms.MessageBox.Show(String.Format("Analysis complete.  Results are in {0}", simulation_output_dir));
        }

        private void SubtleErrSim_Click(object sender, RibbonControlEventArgs e)
        {
            System.Windows.Forms.MessageBox.Show("foo12");

            //// init a RNG
            //var rng = new Random();

            //// get inputs
            //var optinput = RibbonHelper.getExperimentInputs();
            //if (OptTuple.get_IsSome(optinput)) {
            //    var c = optinput.Value.Item1;
            //    var simulation_output_dir = optinput.Value.Item2;

            //    // get sig values
            //    var thresh = 0.05;
            //    Double.TryParse(this.SensitivityTextBox.Text, out thresh);

            //    // show progress bar
            //    var pb = new ProgBar(0, 100);
            //    pb.Show();

            //    // run simulation
            //    RunSubletyExperiment(app, current_workbook, rng, c, simulation_output_dir, thresh, pb);

            //    // close progbar
            //    pb.Close();

            //    // inform user
            //    System.Windows.Forms.MessageBox.Show(String.Format("Analysis complete.  Results are in {0}", simulation_output_dir));
            //}
        }
#endregion BUTTON_HANDLERS

        #region UTILITY_FUNCTIONS
        private static FSharpOption<double> GetSignificance(string input, string label)
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
        #endregion UTILITY_FUNCTIONS
    }
}
