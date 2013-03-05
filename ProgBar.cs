using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Threading;
using System.Text.RegularExpressions;
using DataDebugMethods;
using TreeNode = DataDebugMethods.TreeNode;
using Excel = Microsoft.Office.Interop.Excel;
using TreeDict = System.Collections.Generic.Dictionary<AST.Address, DataDebugMethods.TreeNode>;
using TreeDictPair = System.Collections.Generic.KeyValuePair<AST.Address, DataDebugMethods.TreeNode>;
using System.Collections;


namespace DataDebug
{
    public partial class ProgBar : Form
    {
        private int TRANSPARENT_COLOR_INDEX = -4142;  //-4142 is the transparent default background
        //private bool toolHasNotRun = true; //this is to keep track of whether the tool has already run without having cleared the colorings
        List<TreeNode> originalColorNodes = new List<TreeNode>(); //List for storing the original colors for all nodes
        List<TreeNode> nodelist;        //This is a list holding all the TreeNodes in the Excel file
        double[][][][] impacts_grid; //This is a multi-dimensional array of doubles that will hold each cell's impact on each of the outputs
        bool[][][][] reachable_grid; //This is a multi-dimensional array of bools that will indicate whether a certain output is reachable from a certain cell
        double[][] min_max_delta_outputs; //This keeps the min and max delta for each output; first index indicates the output index; second index 0 is the min delta, 1 is the max delta for that output
        List<TreeNode> ranges;  //This is a list of all the ranges we have identified
        List<TreeNode> charts;  //This is a list of all the charts in the workbook
        List<StartValue> starting_outputs; //This will store the values of all the output nodes at the start of the procedure for swapping values (fuzzing)
        List<TreeNode> output_cells; //This will store all the output nodes at the start of the fuzzing procedure
        List<double[]>[] reachable_impacts_grid;  //This will store impacts for cells reachable from a particular output
        double[][][] reachable_impacts_grid_array; //This will store impacts for cells reachable from a particular output in array form
        int input_cells_in_computation_count = 0;
        int raw_input_cells_in_computation_count = 0;
        private Regex[] regex_array;
        int formula_cells_count;
        TreeDict nodes;
        System.Diagnostics.Stopwatch global_stopwatch = new System.Diagnostics.Stopwatch();
        string stats_text = "";
        TimeSpan swapping_timespan;
        TimeSpan impact_scoring_timespan;
        TimeSpan tree_building_timespan;
        int[][][] times_perturbed;
        int outliers_count;

        public ProgBar(int min, int max)
        {
            this.Visible = true;
            InitializeComponent();
            progressBar1.Minimum = min;
            progressBar1.Maximum = max;
            // Start the BackgroundWorker.
            backgroundWorker1 = new BackgroundWorker();
            backgroundWorker1.DoWork += new DoWorkEventHandler(backgroundWorker1_DoWork);
            backgroundWorker1.ProgressChanged += new ProgressChangedEventHandler(backgroundWorker1_ProgressChanged);
            backgroundWorker1.WorkerReportsProgress = true;
            progressBar1.Value = progressBar1.Minimum;
            backgroundWorker1.RunWorkerAsync();
        }

        private void ProgBar_Load(object sender, System.EventArgs e)
        {
        }

        public void SetProgress(int progress)
        {
            progressBar1.Value = progress;
        }

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            /*
            global_stopwatch.Reset();
            global_stopwatch.Start();
            stats_text = "";

            //Construct a new tree every time the tool is run
            nodelist = new List<TreeNode>();        //This is a list holding all the TreeNodes in the Excel file
            ranges = new List<TreeNode>();        //This is a list holding all the ranges of TreeNodes in the Excel file
            charts = new List<TreeNode>();        //This is a list holding all the chart TreeNodes in the Excel file

            for (int i = 0; i < originalColorNodes.Count; i++)
            {
                if (originalColorNodes[i].getWorkbookObject() == Globals.ThisAddIn.Application.ActiveWorkbook)
                {
                    originalColorNodes.RemoveAt(i);
                }
            }
            foreach (Excel.Worksheet worksheet in Globals.ThisAddIn.Application.Worksheets)
            {
                foreach (Excel.Range cell in worksheet.UsedRange)
                {
                    TreeNode n = new TreeNode(cell.Address, cell.Worksheet, Globals.ThisAddIn.Application.ActiveWorkbook);  //Create a TreeNode for every cell with the name being the cell's address and set the node's worksheet appropriately
                    //n.setOriginalColor(cell.Interior.ColorIndex);
                    n.setOriginalColor(System.Drawing.ColorTranslator.FromOle((int)cell.Interior.Color));
                    originalColorNodes.Add(n);
                }
            }
            constructTree();
            backgroundWorker1.ReportProgress(20);
            perturbationAnalysis();
            backgroundWorker1.ReportProgress(80);
            outlierAnalysis();
            backgroundWorker1.ReportProgress(100);
            showResults();
            */
            while (progressBar1.Value != 100)
            {
                Thread.Sleep(50);
                backgroundWorker1.ReportProgress(progressBar1.Value);
            }
            
            //HERE WE WILL CALL THE FUNCTIONS TO DO THE ANALYSIS
            // 1. CONSTRUCT TREE ==> report progress
            // 2. PERTURBATIONS ==> report progress
            // 3. IMPACT SCORING / LOOK FOR OUTLIERS ==> report progress
            //for (int i = this.progressBar1.Minimum; i <= this.progressBar1.Maximum; i++)
            //{
            //    Thread.Sleep(50);
            //    backgroundWorker1.ReportProgress(i);
            //}
        }

        private void backgroundWorker1_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            progressBar1.Value = e.ProgressPercentage;
            if (progressBar1.Value == progressBar1.Maximum)
            {
                Hide();
            }
        }

        /*
         * This method constructs the dependency graph from the worksheet.
         * It analyzes formulas and looks for references to cells or ranges of cells.
         * It also looks for any charts, and adds those to the dependency graph as well. 
         * After the dependency graph is constructed, we use it to determine and propagate weights to all nodes in the graph. 
         * This method also contains the perturbation procedure and outlier analysis logic.
         * In the end, a text representation of the dependency graph is given in GraphViz format. It includes the entire graph and the weights of the nodes.
         */
        private void constructTree()
        {
            input_cells_in_computation_count = 0;
            raw_input_cells_in_computation_count = 0;

            // Get a range representing the formula cells for each worksheet in each workbook
            ArrayList formulaRanges = ConstructTree.GetFormulaRanges(Globals.ThisAddIn.Application.Worksheets, Globals.ThisAddIn.Application);
            formula_cells_count = ConstructTree.CountFormulaCells(formulaRanges);

            // Create nodes for every cell containing a formula
            // old nodes_grid coordinates were:
            //  1st: worksheet index
            //  2nd: row
            //  3rd: col
            nodes = ConstructTree.CreateFormulaNodes(formulaRanges, Globals.ThisAddIn.Application);

            //This is the list of all ranges referenced in formulas
            List<Excel.Range> referencedRangesList = new List<Excel.Range>();
            //This is the list of TreeNodes for all ranges referenced in formulas
            List<TreeNode> referencedRangesNodeList = new List<TreeNode>();

            int formulaNodesCount = nodes.Count;
            //Now we parse the formulas in nodes to extract any range and cell references
            for (int nodeIndex = 0; nodeIndex < formulaNodesCount; nodeIndex++)
            //foreach (KeyValuePair<AST.Address, TreeNode> nodePair in nodes)
            {
                TreeNode node = nodes.ElementAt(nodeIndex).Value; // nodePair.Value;

                //For each of the ranges found in the formula by the parser, do the following:
                foreach (Excel.Range range in ExcelParserUtility.GetReferencesFromFormula(node.getFormula(), node.getWorkbookObject(), node.getWorksheetObject()))
                {
                    TreeNode rangeNode = null;
                    //See if there is an existing node for this range already in referencedRangesNodeList; if there is, do not add it again - just grab the existing one
                    foreach (TreeNode existingNode in referencedRangesNodeList)
                    {
                        if (existingNode.getName().Equals(range.Address))
                        {
                            rangeNode = existingNode;
                            break;
                        }
                    }
                    if (rangeNode == null)
                    {
                        //TODO CORRECT THE WORKBOOK PARAMETER IN THIS LINE: (IT SHOULD BE THE WORKBOOK OF range, WHICH SHOULD COME FROM GetReferencesFromFormula
                        rangeNode = new TreeNode(range.Address, range.Worksheet, node.getWorkbookObject());
                        referencedRangesList.Add(range);
                        referencedRangesNodeList.Add(rangeNode);
                        ranges.Add(rangeNode);
                    }

                    foreach (Excel.Range cell in range)
                    {
                        TreeNode cellNode = null;
                        //See if there is an existing node for this cell already in nodes; if there is, do not add it again - just grab the existing one
                        if (nodes.TryGetValue(ExcelParser.GetAddress(cell.Address[true, true, Excel.XlReferenceStyle.xlR1C1, false], node.getWorkbookObject(), cell.Worksheet), out cellNode))
                        {
                            //cellNode is set to the value found in the dictionary: no need to do the following:
                            //cellNode = nodes[ExcelParser.GetAddress(cell.Address[true, true, Excel.XlReferenceStyle.xlR1C1, false], node.getWorkbookObject(), cell.Worksheet)];
                        }
                        //if there isn't, create a node and add it to nodes and to referencedRangesList
                        else
                        {
                            //TODO CORRECT THE WORKBOOK PARAMETER IN THIS LINE: (IT SHOULD BE THE WORKBOOK OF cell, WHICH SHOULD COME FROM GetReferencesFromFormula
                            var addr = ExcelParser.GetAddress(cell.Address[true, true, Excel.XlReferenceStyle.xlR1C1, false], node.getWorkbookObject(), cell.Worksheet);
                            cellNode = new TreeNode(cell.Address, cell.Worksheet, node.getWorkbookObject());
                            nodes.Add(addr, cellNode);
                        }
                        rangeNode.addParent(cellNode);
                        cellNode.addChild(node);
                        node.addParent(cellNode);
                    }
                }
            }

            //string tree = "";
            //foreach (KeyValuePair<AST.Address, TreeNode> nodePair in nodes)
            //{   
            //    tree += nodePair.Value.toGVString(0.0) + "\n";
            //}
            //tree = "digraph g{" + tree + "}";
            //Display disp = new Display();
            //disp.textBox1.Text = tree;
            //disp.ShowDialog();

            //TODO -- we are not able to capture ranges that are identified in stored procedures or macros, just ones referenced in formulas
            //TODO -- Dealing with fuzzing of charts -- idea: any cell that feeds into a chart is essentially an output; the chart is just a visual representation (can charts operate on values before they are displayed? don't think so...)
            starting_outputs = new List<StartValue>(); //This will store the values of all the output nodes at the start of the procedure for swapping values (fuzzing)
            output_cells = new List<TreeNode>(); //This will store all the output nodes at the start of the fuzzing procedure

            ConstructTree.StoreOutputs(starting_outputs, output_cells, nodes);

            //Tree building stopwatch
            tree_building_timespan = global_stopwatch.Elapsed;
        }

        private void perturbationAnalysis()
        {
            //Disable screen updating during perturbation to speed things up
            Globals.ThisAddIn.Application.ScreenUpdating = false;

            //Grids for storing influences
            double[][][] influences_grid = null;
            times_perturbed = null;
            //influences_grid and times_perturbed are passed by reference so that they can be modified in the setUpGrids method
            ConstructTree.setUpGrids(ref influences_grid, ref times_perturbed, Globals.ThisAddIn.Application.Worksheets, Globals.ThisAddIn.Application.Charts);

            outliers_count = 0;
            //Procedure for swapping values within ranges, one cell at a time
            List<TreeNode> swap_domain;
            swap_domain = ranges;

            //Initialize min_max_delta_outputs
            min_max_delta_outputs = new double[output_cells.Count][];
            for (int i = 0; i < output_cells.Count; i++)
            {
                min_max_delta_outputs[i] = new double[2];
                min_max_delta_outputs[i][0] = -1.0;
                min_max_delta_outputs[i][1] = 0.0;
            }

            //Initialize impacts_grid 
            //Initialize reachable_grid
            impacts_grid = new double[Globals.ThisAddIn.Application.Worksheets.Count][][][];
            reachable_grid = new bool[Globals.ThisAddIn.Application.Worksheets.Count][][][];
            foreach (Excel.Worksheet worksheet in Globals.ThisAddIn.Application.Worksheets)
            {
                impacts_grid[worksheet.Index - 1] = new double[worksheet.UsedRange.Rows.Count + worksheet.UsedRange.Row][][];
                reachable_grid[worksheet.Index - 1] = new bool[worksheet.UsedRange.Rows.Count + worksheet.UsedRange.Row][][];
                for (int row = 0; row < (worksheet.UsedRange.Rows.Count + worksheet.UsedRange.Row); row++)
                {
                    impacts_grid[worksheet.Index - 1][row] = new double[worksheet.UsedRange.Columns.Count + worksheet.UsedRange.Column][];
                    reachable_grid[worksheet.Index - 1][row] = new bool[worksheet.UsedRange.Columns.Count + worksheet.UsedRange.Column][];
                    for (int col = 0; col < (worksheet.UsedRange.Columns.Count + worksheet.UsedRange.Column); col++)
                    {
                        impacts_grid[worksheet.Index - 1][row][col] = new double[output_cells.Count];
                        reachable_grid[worksheet.Index - 1][row][col] = new bool[output_cells.Count];
                        for (int i = 0; i < output_cells.Count; i++)
                        {
                            impacts_grid[worksheet.Index - 1][row][col][i] = 0.0;
                            reachable_grid[worksheet.Index - 1][row][col][i] = false;
                        }
                    }
                }
            }

            //Initialize reachable_impacts_grid
            reachable_impacts_grid = new List<double[]>[output_cells.Count];
            for (int i = 0; i < output_cells.Count; i++)
            {
                reachable_impacts_grid[i] = new List<double[]>();
            }

            //Propagate weights  -- find the weights of all outputs and set up the reachable_grid entries
            foreach (TreeDictPair tdp in nodes)
            {
                var node = tdp.Value;
                if (!node.hasParents())
                {
                    node.setWeight(1.0);  //Set the weight of all input nodes to 1.0 to start
                    //Now we propagate it's weight to all of it's children
                    TreeNode.propagateWeightUp(node, 1.0, node, output_cells, reachable_grid, reachable_impacts_grid);
                    raw_input_cells_in_computation_count++;
                }
            }

            //Convert reachable_impacts_grid to array form
            reachable_impacts_grid_array = new double[output_cells.Count][][];
            for (int i = 0; i < output_cells.Count; i++)
            {
                reachable_impacts_grid_array[i] = reachable_impacts_grid[i].ToArray();
            }

            ConstructTree.SwappingProcedure(swap_domain, ref input_cells_in_computation_count, ref min_max_delta_outputs, ref impacts_grid, ref times_perturbed, ref output_cells, ref reachable_grid, ref starting_outputs, ref reachable_impacts_grid_array);
            
            //Stop timing swapping procedure:
            swapping_timespan = global_stopwatch.Elapsed;
        }

        private void outlierAnalysis()
        {
            ConstructTree.ComputeZScoresAndFindOutliers(output_cells, reachable_impacts_grid_array, impacts_grid, times_perturbed, Globals.ThisAddIn.Application.Worksheets, outliers_count);
            //Stop timing the zscore computation and outlier finding
            impact_scoring_timespan = global_stopwatch.Elapsed;

            Globals.ThisAddIn.Application.ScreenUpdating = true;

            
        }
        private void showResults()
        {
            // Format and display the TimeSpan value. 
            string tree_building_time = tree_building_timespan.TotalSeconds + ""; //String.Format("{0:00}:{1:00}.{2:00}", tree_building_timespan.Minutes, tree_building_timespan.Seconds, tree_building_timespan.Milliseconds / 10);
            string swapping_time = (swapping_timespan.TotalSeconds - tree_building_timespan.TotalSeconds) + ""; //String.Format("{0:00}:{1:00}.{2:00}", swapping_timespan.Minutes, swapping_timespan.Seconds, swapping_timespan.Milliseconds / 10);
            string impact_scoring_time = (impact_scoring_timespan.TotalSeconds - swapping_timespan.TotalSeconds) + ""; //String.Format("{0:00}:{1:00}.{2:00}", z_score_timespan.Minutes, z_score_timespan.Seconds, z_score_timespan.Milliseconds / 10);
            global_stopwatch.Stop();
            // Get the elapsed time as a TimeSpan value.
            TimeSpan global_timespan = global_stopwatch.Elapsed;
            //string global_time = String.Format("{0:00}:{1:00}:{2:00}.{3:00}", global_timespan.Hours, global_timespan.Minutes, global_timespan.Seconds, global_timespan.Milliseconds / 10);
            string global_time = global_timespan.TotalSeconds + ""; //(tree_building_timespan.TotalSeconds + swapping_timespan.TotalSeconds + z_score_timespan.TotalSeconds + average_z_score_timespan.TotalSeconds + outlier_detection_timespan.TotalSeconds + outlier_coloring_timespan.TotalSeconds) + ""; //String.Format("{0:00}:{1:00}.{2:00}",

            Display timeDisplay = new Display();
            stats_text += "" //+ "Benchmark:\tNumber of formulas:\tRaw input count:\tInputs to computations:\tTotal (s):\tTree Construction (s):\tSwapping (s):\tZ-Score Calculation (s):\t"
                //  + "Outlier Detection (s):\tOutlier Coloring (s):\t"
                //+ "Outliers found:\n"
                //"Formula cells:\t" + formula_cells_count + "\n"
                //+ "Number of input cells involved in computations:\t" + input_cells_in_computation_count
                //+ "\nExecution times (seconds): "
                + Globals.ThisAddIn.Application.ActiveWorkbook.Name + "\t"
                + formula_cells_count + "\t"
                + raw_input_cells_in_computation_count + "\t"
                + input_cells_in_computation_count + "\t"
                + global_time + "\t"
                + tree_building_time + "\t"
                + swapping_time + "\t"
                + impact_scoring_time + "\t"
                + outliers_count;
            timeDisplay.textBox1.Text = stats_text;
            timeDisplay.ShowDialog();
        }
    }
}
