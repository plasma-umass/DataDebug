using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.Office.Tools.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using System.Text.RegularExpressions;
using DataDebugMethods;
using Microsoft.FSharp.Core;
using TreeNode = DataDebugMethods.TreeNode;
using System.Diagnostics;
using TreeDict = System.Collections.Generic.Dictionary<AST.Address, DataDebugMethods.TreeNode>;
using TreeDictPair = System.Collections.Generic.KeyValuePair<AST.Address, DataDebugMethods.TreeNode>;

namespace DataDebug
{
    public partial class Ribbon
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
        System.Diagnostics.Stopwatch global_stopwatch = new System.Diagnostics.Stopwatch();
        string stats_text = "";

        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        // TODO move analysis into separate lib

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
            TimeSpan impact_scoring_timespan = global_stopwatch.Elapsed;
            TimeSpan swapping_timespan = global_stopwatch.Elapsed;
            input_cells_in_computation_count = 0;
            raw_input_cells_in_computation_count = 0;

            // Get a range representing the formula cells for each worksheet in each workbook
            ArrayList analysisRanges = ConstructTree.GetFormulaRanges(Globals.ThisAddIn.Application.Worksheets, Globals.ThisAddIn.Application);
            formula_cells_count = ConstructTree.CountFormulaCells(analysisRanges);

            // Create nodes for every cell containing a formula
            // old nodes_grid coordinates were:
            //  1st: worksheet index
            //  2nd: row
            //  3rd: col
            TreeDict nodes = ConstructTree.CreateFormulaNodes(analysisRanges, Globals.ThisAddIn.Application);
            
            // Get the names of all worksheets in the workbook and store them in the array worksheet_names
            String[] worksheet_names = new String[Globals.ThisAddIn.Application.Worksheets.Count];
            // ditto, but for the actual references
            Excel.Worksheet[] worksheet_refs = new Excel.Worksheet[Globals.ThisAddIn.Application.Worksheets.Count]; 

            int index_worksheet_names = 0; // Index for populating the worksheet_names
            foreach (Excel.Worksheet worksheet in Globals.ThisAddIn.Application.Worksheets)
            {
                worksheet_names[index_worksheet_names] = worksheet.Name;
                worksheet_refs[index_worksheet_names] = worksheet;
                index_worksheet_names++;
            }

            // worksheet_range is guaranteed to consist only of a collection of formula cells
            foreach (Excel.Range worksheet_range in analysisRanges)
            {
                foreach (Excel.Range c in worksheet_range)
                {
                    // Sanity check-- ensure that cells from ranges in analysisRanges are inside the UsedRange
                    Debug.Assert(DataDebugMethods.Utility.InsideUsedRange(c, Globals.ThisAddIn.Application.ActiveWorkbook), "This spreadsheet violates a condition thought to be impossible.");
                    TreeNode formula_cell;
                    AST.Address addr = Utility.ParseXLAddress(c, Globals.ThisAddIn.Application.ActiveWorkbook);
                    if (!nodes.TryGetValue(addr, out formula_cell))
                    {
                        throw new Exception("Sometimes you eat the bear, and sometimes, well, he eats you.");
                    }

                    string formula = c.Formula;  //The formula contained in the cell
                    ConstructTree.StripLookups(formula);

                    MatchCollection matchedRanges = null;
                    MatchCollection matchedCells = null;
                    int ws_index = 1;   //the index of the worksheet in which the formula cell resides
                    
                    for (int i = 0; i < worksheet_names.Count(); i++)
                    {
                        string s = worksheet_names[i];  //the name of the worksheet that may be referenced in the formula
                        Excel.Worksheet ws_ref = worksheet_refs[i];
                        string worksheet_name = s.Replace("+", @"\+").Replace("^", @"\^").Replace("$", @"\$").Replace(".", @"\."); //Escape certain characters in the regular expression
                        //First look for range references of the form 'worksheet_name'!A1:A10 in the formula (with quotation marks around the name)
                        ConstructTree.FindRangeReferencesWithQuotes(ref formula, worksheet_name, matchedRanges, regex_array, ws_index, ranges, formula_cell, ws_ref, Globals.ThisAddIn.Application.ActiveWorkbook, Globals.ThisAddIn.Application.Worksheets[ws_index], nodes);

                        //Next look for range references of the form worksheet_name!A1:A10 in the formula (no quotation marks around the name)
                        ConstructTree.FindRangeReferencesWithoutQuotes(ref formula, worksheet_name, matchedRanges, regex_array, ws_index, ranges, formula_cell, ws_ref, Globals.ThisAddIn.Application.ActiveWorkbook, Globals.ThisAddIn.Application.Worksheets[ws_index], nodes);

                        // Now we look for references of the kind 'worksheet_name'!A1 (with quotation marks)
                        ConstructTree.FindCellReferencesWithQuotes(ref formula, worksheet_name, matchedCells, regex_array, ws_index, ranges, formula_cell, ws_ref, Globals.ThisAddIn.Application.ActiveWorkbook, Globals.ThisAddIn.Application.Worksheets, nodes);
                        
                        ws_index++;
                    }

                    ConstructTree.FindRangeReferencesInCurrentWorksheet(ref formula, matchedRanges, matchedCells, regex_array, ws_index, ranges, formula_cell, Globals.ThisAddIn.Application.ActiveWorkbook, Globals.ThisAddIn.Application.Worksheets, nodes, c);

                    ConstructTree.FindNamedRangeReferences(ref formula, matchedRanges, matchedCells, regex_array, ws_index, ranges, formula_cell, Globals.ThisAddIn.Application.ActiveWorkbook, Globals.ThisAddIn.Application.Worksheets, nodes, c, Globals.ThisAddIn.Application.Names);

                    ConstructTree.FindCellReferencesInCurrentWorksheet(ref formula, matchedRanges, matchedCells, regex_array, ws_index, ranges, formula_cell, Globals.ThisAddIn.Application.ActiveWorkbook, Globals.ThisAddIn.Application.Worksheets, nodes, c);
                     
                }
            }

            //Display textbox with GraphViz representation of the dependence graph
            //Display disp = new Display();
            //disp.textBox1.Text = ConstructTree.GenerateGraphVizTree(nodes);
            //disp.ShowDialog();

            ConstructTree.FindReferencesInCharts(regex_array, ranges, Globals.ThisAddIn.Application.ActiveWorkbook, Globals.ThisAddIn.Application.Charts, nodes, worksheet_names, worksheet_refs, Globals.ThisAddIn.Application.Worksheets, Globals.ThisAddIn.Application.Names, nodelist);

            //TODO -- we are not able to capture ranges that are identified in stored procedures or macros, just ones referenced in formulas
            //TODO -- Dealing with fuzzing of charts -- idea: any cell that feeds into a chart is essentially an output; the chart is just a visual representation (can charts operate on values before they are displayed? don't think so...)
            starting_outputs = new List<StartValue>(); //This will store the values of all the output nodes at the start of the procedure for swapping values (fuzzing)
            output_cells = new List<TreeNode>(); //This will store all the output nodes at the start of the fuzzing procedure

            ConstructTree.StoreOutputs(starting_outputs, output_cells, nodes);
         
            //Tree building stopwatch
            TimeSpan tree_building_timespan = global_stopwatch.Elapsed;

            //Disable screen updating during perturbation to speed things up
            Globals.ThisAddIn.Application.ScreenUpdating = false;

            //Grids for storing influences
            double[][][] influences_grid = null;
            int[][][] times_perturbed = null;
            //influences_grid and times_perturbed are passed by reference so that they can be modified in the setUpGrids method
            ConstructTree.setUpGrids(ref influences_grid, ref times_perturbed, Globals.ThisAddIn.Application.Worksheets, Globals.ThisAddIn.Application.Charts);
            
            int outliers_count = 0; 
            //Procedure for swapping values within ranges, one cell at a time
            if (!checkBox2.Checked) //Checks if the option for swapping values simultaneously is checked (not checked by default)
            {
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
                //System.Windows.Forms.MessageBox.Show("There are " + nodes.Count.ToString() + " nodes in our dictionary.");
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

                ConstructTree.SwappingProcedure(swap_domain, input_cells_in_computation_count, min_max_delta_outputs, impacts_grid, times_perturbed, output_cells, reachable_grid, starting_outputs, ref reachable_impacts_grid_array, reachable_impacts_grid);

                //Stop timing swapping procedure:
                swapping_timespan = global_stopwatch.Elapsed;

                ConstructTree.ComputeZScoresAndFindOutliers(output_cells, reachable_impacts_grid_array, impacts_grid, times_perturbed, Globals.ThisAddIn.Application.Worksheets, outliers_count);
                //Stop timing the zscore computation and outlier finding
                impact_scoring_timespan = global_stopwatch.Elapsed;
            }

            Globals.ThisAddIn.Application.ScreenUpdating = true;
            
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

            //Procedure for swapping values within ranges, replacing all repeated values at once
            if (checkBox2.Checked) //Checks if the option for swapping values simultaneously is checked
            {
                List<TreeNode> swap_domain;
                swap_domain = ranges;
                
                foreach (TreeNode node in swap_domain)
                {
                    bool all_children_are_charts = true;
                    if (node.isRange() && node.hasChildren())
                    {
                        //bool children_are_charts = false;
                        foreach (TreeNode child in node.getChildren())
                        {
                            if (!child.isChart())
                            {
                                all_children_are_charts = false;
                            }
                        }
                    }
                    //For each range node, do the following:
                    if (node.isRange() && !all_children_are_charts)
                    {
                        double[] influences = new double[node.getParents().Count];  //Array to keep track of the influence values for every cell
                        int influence_index = 0;        //Keeps track of the current position in the influences array
                        double max_total_delta = 0;     //The maximum influence found (for normalizing)
                        double min_total_delta =-1;     //The minimum influence found (for normalizing)
                        //Swapping values; loop over all nodes in the range
                        foreach (TreeNode parent in node.getParents())
                        {
                            String twin_cells_string = parent.getName();
                            //Find any nodes with a matching value and keep track of them
                            int twin_count = 1;     //This will keep track of the number of cells that have this exact value
                            foreach (TreeNode twin in node.getParents())
                            {
                                if (twin.getName() == parent.getName()) // if twin is the same cell as the current cell being examined, don't do anything
                                {
                                    continue;
                                }
                                if (twin.getWorksheetObject().get_Range(twin.getName()).Value == parent.getWorksheetObject().get_Range(parent.getName()).Value)
                                {
                                    twin_cells_string = twin_cells_string + "," + twin.getName();
                                    twin_count++;
                                }
                            }
                            Excel.Range twin_cells = parent.getWorksheetObject().get_Range(twin_cells_string);
                            String[] formulas = new String[twin_count]; //Stores the formulas in the twin_cells
                            int i = 0; //Counter for indexing within the formulas array
                            foreach (Excel.Range cell in twin_cells)
                            {
                                if (cell.HasFormula)
                                {
                                    formulas[i] = cell.Formula;
                                }
                                i++;
                            }
                            double start_value = parent.getWorksheetObject().get_Range(parent.getName()).Value;
                            double total_delta = 0;
                            double delta = 0;
                            foreach (TreeNode sibling in node.getParents())
                            {
                                if (sibling.getName() == parent.getName())
                                {
                                    continue;
                                }
                                Excel.Range sibling_cell = sibling.getWorksheetObject().get_Range(sibling.getName());
                                twin_cells.Value = sibling_cell.Value;
                                int index = 0;
                                delta = 0;
                                foreach (TreeNode n in output_cells)
                                {
                                    if (starting_outputs[index].get_double() != 0)
                                    {
                                        delta = Math.Abs(starting_outputs[index].get_double() - n.getWorksheetObject().get_Range(n.getName()).Value) / Math.Abs(starting_outputs[index].get_double());
                                    }
                                    else
                                    {
                                        delta = Math.Abs(starting_outputs[index].get_double() - n.getWorksheetObject().get_Range(n.getName()).Value);
                                    }
                                    index++;
                                    total_delta = total_delta + delta;
                                }
                            }
                            total_delta = total_delta / (node.getParents().Count - 1);
                            total_delta = total_delta / twin_count;
                            influences[influence_index] = total_delta;
                            influence_index++;
                            if (max_total_delta < total_delta)
                            {
                                max_total_delta = total_delta;
                            }
                            if (min_total_delta > total_delta || min_total_delta == -1)
                            {
                                min_total_delta = total_delta;
                            }
                            twin_cells.Value = start_value;
                            int j = 0;
                            foreach (Excel.Range cell in twin_cells)
                            {
                                if (formulas[j] != null)
                                    cell.Formula = formulas[j];
                                j++;
                            }
                            twin_cells.Interior.Color = System.Drawing.Color.Beige;
                        }
                        int ind = 0;
                        foreach (TreeNode parent in node.getParents())
                        {
                            if (max_total_delta != 0)
                            {
                                influences[ind] = (influences[ind] - min_total_delta) / max_total_delta;
                            }
                            ind++;
                        }
                        int indexer = 0;
                        foreach (TreeNode parent in node.getParents())
                        {
                            Excel.Range cell = parent.getWorksheetObject().get_Range(parent.getName());
                            cell.Interior.Color = System.Drawing.Color.FromArgb(Convert.ToInt32(255 - influences[indexer] * 255), 255, 255);
                            indexer++;
                        }
                    }
                }
            }
        }
        
        //Action for "Analyze Worksheet" button
        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            global_stopwatch.Reset();
            global_stopwatch.Start();
            stats_text = "";
            
            //Construct a new tree every time the tool is run
            nodelist = new List<TreeNode>();        //This is a list holding all the TreeNodes in the Excel file
            ranges = new List<TreeNode>();        //This is a list holding all the ranges of TreeNodes in the Excel file
            charts = new List<TreeNode>();        //This is a list holding all the chart TreeNodes in the Excel file
            
            //Compile regular expressions
            if (toggle_compile_regex.Checked)
            {
                regex_array = new Regex[Globals.ThisAddIn.Application.Worksheets.Count * 4 + 2];
                int worksheet_index = 0;
                foreach (Excel.Worksheet worksheet in Globals.ThisAddIn.Application.Worksheets)
                {
                    string worksheet_name = worksheet.Name.Replace("+", @"\+").Replace("^", @"\^").Replace("$", @"\$").Replace(".", @"\."); //Escape certain characters in the regular expression
                    regex_array[worksheet_index*4] = new Regex(@"('" + worksheet_name + @"'!\$?[A-Z]+\$?[1-9]\d*:\$?[A-Z]+\$?[1-9]\d*)", RegexOptions.Compiled);
                    regex_array[worksheet_index*4 + 1] = new Regex(@"(" + worksheet_name + @"!\$?[A-Z]+\$?[1-9]\d*:\$?[A-Z]+\$?[1-9]\d*)", RegexOptions.Compiled);
                    regex_array[worksheet_index*4 + 2] = new Regex(@"('" + worksheet_name + @"'!\$?[A-Z]+\$?[1-9]\d*)", RegexOptions.Compiled);
                    regex_array[worksheet_index*4 + 3] = new Regex(@"(" + worksheet_name + @"!\$?[A-Z]+\$?[1-9]\d*)", RegexOptions.Compiled);
                    worksheet_index++;
                }
                regex_array[regex_array.Length - 2] = new Regex(@"(\$?[A-Z]+\$?[1-9]\d*:\$?[A-Z]+\$?[1-9]\d*)", RegexOptions.Compiled);
                regex_array[regex_array.Length - 1] = new Regex(@"(\$?[A-Z]+\$?[1-9]\d*)", RegexOptions.Compiled);
            }

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
        }

        //Button for testing random code :)
        private void button7_Click(object sender, RibbonControlEventArgs e)
        {
            //System.Windows.Forms.MessageBox.Show(Globals.ThisAddIn.Application.ActiveWorkbook.Path + "");
            //System.Windows.Forms.MessageBox.Show(Globals.ThisAddIn.Application.Workbooks[1] + "");
            //foreach (Excel.Chart chart in Globals.ThisAddIn.Application.Charts)
            //{
            //    foreach (Excel.Series series in (Excel.SeriesCollection)chart.SeriesCollection(Type.Missing))
            //    {
            //        string formula = series.Formula;  //The formula contained in the cell
            //        System.Windows.Forms.MessageBox.Show(formula);
            //    }
            //}
            //Form1 form = new Form1();
            //form.Visible = true; // Show();
            ProgBar pb = new ProgBar(0, 100);
            pb.Show();
        }

        //Action for "Clear coloring" button
        private void button8_Click(object sender, RibbonControlEventArgs e)
        {
            for (int i = 0; i < originalColorNodes.Count; i++)
            {
                if (originalColorNodes[i].getWorkbookObject() == Globals.ThisAddIn.Application.ActiveWorkbook)
                {
                    //If the node is just a cell, clear any coloring
                    if (!originalColorNodes[i].isChart() && !originalColorNodes[i].isRange())
                    {
                        //System.Windows.Forms.MessageBox.Show(node.getName() + " " + node.getOriginalColor());
                        //node.getWorksheetObject().get_Range(node.getName()).Interior.ColorIndex = 0;
                        //node.getWorksheetObject().get_Range(node.getName()).Interior.ColorIndex = node.getOriginalColor();
                        if (!(originalColorNodes[i].getOriginalColor() + "").Equals("Color [White]"))
                        {
                            originalColorNodes[i].getWorksheetObject().get_Range(originalColorNodes[i].getName()).Interior.Color = originalColorNodes[i].getOriginalColor();
                        }
                        else
                        {
                            originalColorNodes[i].getWorksheetObject().get_Range(originalColorNodes[i].getName()).Interior.ColorIndex = TRANSPARENT_COLOR_INDEX;  //-4142 is the transparent default background for cells
                        }
                    }
                    originalColorNodes.RemoveAt(i);
                    i--;
                }
            }
        }

        private void TestParser_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;
            Excel.Worksheet ws = Globals.ThisAddIn.Application.ActiveSheet;
            Excel.Range fn_cell = Globals.ThisAddIn.Application.ActiveCell;
            if (fn_cell.HasFormula)
            {
                string formula = Convert.ToString(fn_cell.Formula);
                IEnumerable<Excel.Range> ranges = null; //ExcelParserUtility.GetReferencesFromFormula(formula, wb, ws);
                foreach (Excel.Range range in ranges)
                {
                    foreach (Excel.Range cell in range)
                    {
                        cell.Interior.Color = System.Drawing.Color.Red;
                    }
                }
            }
            else
            {
                System.Windows.Forms.MessageBox.Show("The currently selected cell is not a formula.");
            }
            
        }
    }
}
