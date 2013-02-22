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
        double[][] min_max_delta_outputs; //This keeps the min and max delta for each output
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
                    Debug.Assert(DataDebugMethods.Utility.InsideUsedRange(c), "This spreadsheet violates a condition thought to be impossible.");
                    TreeNode formula_cell;
                    AST.Address addr = Utility.ParseXLAddress(c);
                    if (!nodes.TryGetValue(addr, out formula_cell))
                    {
                        throw new Exception("Sometimes you eat the bear, and sometimes, well, he eats you.");
                    }

                    string formula = c.Formula;  //The formula contained in the cell
                    ConstructTree.StripLookups(formula);

                    MatchCollection matchedRanges = null;
                    MatchCollection matchedCells = null;
                    int ws_index = 1;   //the index of the worksheet in which the formula cell resides
                    //foreach (string s in worksheet_names)

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
            influences_grid = new double[Globals.ThisAddIn.Application.Worksheets.Count + Globals.ThisAddIn.Application.Charts.Count][][];
            times_perturbed = new int[Globals.ThisAddIn.Application.Worksheets.Count + Globals.ThisAddIn.Application.Charts.Count][][];
            foreach (Excel.Worksheet worksheet in Globals.ThisAddIn.Application.Worksheets)
            {
                influences_grid[worksheet.Index - 1] = new double[worksheet.UsedRange.Rows.Count + worksheet.UsedRange.Row][];
                times_perturbed[worksheet.Index - 1] = new int[worksheet.UsedRange.Rows.Count + worksheet.UsedRange.Row][];
                for (int row = 0; row < (worksheet.UsedRange.Rows.Count + worksheet.UsedRange.Row); row++)
                {
                    influences_grid[worksheet.Index - 1][row] = new double[worksheet.UsedRange.Columns.Count + worksheet.UsedRange.Column];
                    times_perturbed[worksheet.Index - 1][row] = new int[worksheet.UsedRange.Columns.Count + worksheet.UsedRange.Column];
                    for (int col = 0; col < (worksheet.UsedRange.Columns.Count + worksheet.UsedRange.Column); col++)
                    {
                        influences_grid[worksheet.Index - 1][row][col] = 0.0;
                        times_perturbed[worksheet.Index - 1][row][col] = 0;
                    }
                }
            }
            
            int outliers_count = 0; 
            //Procedure for swapping values within ranges, one cell at a time
            if (!checkBox2.Checked) //Checks if the option for swapping values simultaneously is checked
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
                            //System.Windows.Forms.MessageBox.Show("output cells count = " + output_cells.Count);
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
                System.Windows.Forms.MessageBox.Show("There are " + nodes.Count.ToString() + " nodes in our dictionary.");
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
                

                foreach (TreeNode range_node in swap_domain)
                {
                    //If this node is not a range, continue to the next node because no perturbations can be done on this node.
                    if (!range_node.isRange())
                    {
                        continue;
                    }
                    //For every range node
                    double[] influences = new double[range_node.getParents().Count]; //Array to keep track of the influence values for every cell in the range
                    //double max_total_delta = 0;     //The maximum influence found (for normalizing)
                    //double min_total_delta = -1;     //The minimum influence found (for normalizing)
                    int swaps_per_range = 30; // 30;
                    if (range_node.getParents().Count <= swaps_per_range)
                    {
                        swaps_per_range = range_node.getParents().Count - 1;
                    }
                    Random rand = new Random();
                    int current_index = 0; 
                    //Swapping values; loop over all nodes in the range
                    foreach (TreeNode parent in range_node.getParents())
                    {
                        //Do not perturb nodes which are intermediate computations
                        if (parent.hasParents())
                        {
                            continue;
                        }
                        input_cells_in_computation_count++;
                        
                        //Generate 30 random indices for swapping with siblings
                        int[] indices = new int[swaps_per_range];
                        if (swaps_per_range == 30)
                        {
                            int indices_pointer = 0;
                            int random_index = 0;
                            for (int i = 0; i < swaps_per_range; i++)
                            {
                                do
                                {
                                    random_index = rand.Next(range_node.getParents().Count);
                                } while (random_index == current_index);
                                indices[indices_pointer] = random_index;
                                indices_pointer++;
                            }
                        }
                        //Generate indices for swapping with siblings -- include all indices but the one for the current node (so as to not swap with itself)
                        else
                        {
                            int sibling_ind = 0; //sibling_ind is a counter that goes through all indices
                            for (int i = 0; i < swaps_per_range; i++)
                            {
                                if (sibling_ind == current_index)       //if the sibling index is the same as the current node's index, go to the next index
                                {
                                    sibling_ind++;
                                }
                                indices[i] = sibling_ind;
                                sibling_ind++;
                            }
                        }

                        Excel.Range cell = parent.getWorksheetObject().get_Range(parent.getName());
                        string formula = "";
                        if (cell.HasFormula)
                        {
                            formula = cell.Formula;
                        }
                        StartValue start_value = new StartValue(cell.Value); //Store the initial value of this cell before swapping
                        double total_delta = 0.0; // Stores the total change in outputs caused by this cell after swapping with every other value in the range
                        double delta = 0.0;   // Stores the change in outputs caused by a single swap
                        //Swapping loop - swap every sibling or a reduced number of siblings (approximately equal to swaps_per_range), for reduced complexity and runtime
                        //foreach (TreeNode sibling in node.getParents())
                        foreach (int sibling_index in indices)
                        {
                            TreeNode sibling = range_node.getParents()[sibling_index];
                            if (sibling.getName() == parent.getName() && sibling.getWorksheetObject() == parent.getWorksheetObject())
                            {
                                continue; // If this is the current cell, move on to the next cell -- this should never happen because the sibling index should never equal the current index
                            }

                            try
                            {
                                times_perturbed[cell.Worksheet.Index - 1][cell.Row - 1][cell.Column - 1]++;
                            }
                            catch
                            {
                                cell.Interior.Color = System.Drawing.Color.Purple;
                            }
                            Excel.Range sibling_cell = sibling.getWorksheetObject().get_Range(sibling.getName());
                            cell.Value = sibling_cell.Value; //This is the swap -- we assign the value of the sibling cell to the value of our cell
                            delta = 0.0;
                            //foreach (TreeNode n in output_cells)
                            for (int i = 0; i < output_cells.Count; i++)
                            {
                                try
                                {
                                    //If this output is not reachable from this cell, continue
                                    if (reachable_grid[cell.Worksheet.Index - 1][cell.Row - 1][cell.Column - 1][i] == false)
                                    {
                                        continue;
                                    }
                                }
                                catch
                                {
                                    continue;
                                }
                                TreeNode n = output_cells[i];
                                if (starting_outputs[i].get_string() == null) // If the output is not a string
                                {
                                    if (!n.isChart())   //If the output is not a chart, it must be a number
                                    {
                                        delta = Math.Abs(starting_outputs[i].get_double() - (double)n.getWorksheetObject().get_Range(n.getName()).Value);  //Compute the absolute change caused by the swap
                                    }
                                    else  // The node is a chart
                                    {
                                        double sum = 0.0;
                                        TreeNode parent_range = n.getParents()[0];
                                        foreach (TreeNode par in parent_range.getParents())
                                        {
                                            sum = sum + (double)par.getWorksheetObject().get_Range(par.getName()).Value;
                                        }
                                        double average = sum / parent_range.getParents().Count;
                                        delta = Math.Abs(starting_outputs[i].get_double() - average);
                                    }
                                }
                                else  // If the output is a string
                                {
                                    if (String.Equals(starting_outputs[i].get_string(), n.getWorksheetObject().get_Range(n.getName()).Value, StringComparison.Ordinal))
                                    {
                                        delta = 0.0;
                                    }
                                    else
                                    {
                                        delta = 1.0;
                                    }
                                }
                                //Add to the impact of the cell for this output
                                impacts_grid[cell.Worksheet.Index - 1][cell.Row - 1][cell.Column - 1][i] += delta;
                                //Compare the min/max values for this output to this delta
                                if (min_max_delta_outputs[i][0] == -1.0)
                                {
                                    min_max_delta_outputs[i][0] = delta;
                                }
                                else
                                {
                                    if (min_max_delta_outputs[i][0] > delta)
                                    {
                                        min_max_delta_outputs[i][0] = delta;
                                    }
                                }
                                if (min_max_delta_outputs[i][1] < delta)
                                {
                                    min_max_delta_outputs[i][1] = delta;
                                }
                                //index++;
                                total_delta = total_delta + delta;
                            }
                        }
                        
                        if (start_value.get_string() == null)
                        {
                            cell.Value = start_value.get_double();
                        }
                        else
                        {
                            cell.Value = start_value.get_string();
                        }
                        if (formula != "")
                        {
                            cell.Formula = formula;
                        }
                        current_index++;
                    }

                    //TODO: if there are overflow issues consider making total_influence an array of doubles (of size 100 for instance) and use each slot as a bin for parts of the sum
                    //each part can be divided by the denominator and then the average_influence is the sum of the entries in the array
                }

                //Convert reachable_impacts_grid to array form
                reachable_impacts_grid_array = new double[output_cells.Count][][];
                for (int i = 0; i < output_cells.Count; i++)
                {
                    reachable_impacts_grid_array[i] = reachable_impacts_grid[i].ToArray(); 
                }

                //Populate reachable_impacts_grid_array from impacts_grid
                for (int i = 0; i < output_cells.Count; i++)
                {
                    for (int d = 0; d < reachable_impacts_grid_array[i].Length; d++)
                    {
                        reachable_impacts_grid_array[i][d] = new double[4] { reachable_impacts_grid_array[i][d][0], 
                            reachable_impacts_grid_array[i][d][1], 
                            reachable_impacts_grid_array[i][d][2], 
                            impacts_grid[(int)reachable_impacts_grid_array[i][d][0]][(int)reachable_impacts_grid_array[i][d][1]][(int)reachable_impacts_grid_array[i][d][2]][i] };
                    }
                }

                //Stop timing swapping procedure:
                swapping_timespan = global_stopwatch.Elapsed;
                
                //Now for each output, compute the z-score of the impact of each input
                for (int i = 0; i < output_cells.Count; i++)
                {
                    //Find the mean for the output
                    double output_sum = 0.0;
                    
                    for (int d = 0; d < reachable_impacts_grid_array[i].Length; d++)
                    {
                        int worksheet_ind = (int)reachable_impacts_grid_array[i][d][0];
                        int row = (int)reachable_impacts_grid_array[i][d][1];
                        int col = (int)reachable_impacts_grid_array[i][d][2];
                        if (times_perturbed[worksheet_ind][row][col] != 0)
                        {
                            output_sum += impacts_grid[worksheet_ind][row][col][i];
                        }
                    }

                    double output_average = 0.0;
                    if (reachable_impacts_grid_array[i].Length != 0)
                    {
                        output_average = output_sum / (double)reachable_impacts_grid_array[i].Length;
                    }
                    else  //if none of the entries can reach this output, all impacts must be equal to 0.
                    {
                        output_average = 0.0;
                    }
                    //Find the sample standard deviation for this output
                    double variance = 0.0;
                    
                    for (int d = 0; d < reachable_impacts_grid_array[i].Length; d++)
                    {
                        int worksheet_ind = (int)reachable_impacts_grid_array[i][d][0];
                        int row = (int)reachable_impacts_grid_array[i][d][1];
                        int col = (int)reachable_impacts_grid_array[i][d][2];
                        if (times_perturbed[worksheet_ind][row][col] != 0)
                        {
                            variance += Math.Pow(output_average - impacts_grid[worksheet_ind][row][col][i], 2) / (double)reachable_impacts_grid_array[i].Length;
                        }
                    }
                    double std_dev = Math.Sqrt(variance);
                    
                    for (int d = 0; d < reachable_impacts_grid_array[i].Length; d++)
                    {
                        int worksheet_ind = (int)reachable_impacts_grid_array[i][d][0];
                        int row = (int)reachable_impacts_grid_array[i][d][1];
                        int col = (int)reachable_impacts_grid_array[i][d][2];
                        if (times_perturbed[worksheet_ind][row][col] != 0)
                        {
                            if (std_dev != 0.0)
                            {
                                reachable_impacts_grid_array[i][d][3] = Math.Abs((impacts_grid[worksheet_ind][row][col][i] - output_average) / std_dev);
                            }
                            else  //std_dev == 0.0
                            {
                                //If the standard deviation is zero, then all the impacts were the same and we shouldn't flag any entries, so set their z-scores to 0.0
                                reachable_impacts_grid_array[i][d][3] = 0.0;
                            }
                        }
                    }
                }

                //Repopulate impacts_grid with the z-scores from reachable_impacts_grid_array
                foreach (Excel.Worksheet worksheet in Globals.ThisAddIn.Application.Worksheets)
                {
                    Excel.Range used_range = worksheet.get_Range("A1");
                    for (int row = 0; row < (worksheet.UsedRange.Rows.Count + worksheet.UsedRange.Row); row++)
                    {
                        for (int col = 0; col < (worksheet.UsedRange.Columns.Count + worksheet.UsedRange.Column); col++)
                        {
                            for (int i = 0; i < output_cells.Count; i++)
                            {
                                impacts_grid[worksheet.Index - 1][row][col][i] = 0.0;
                            }
                        }
                    }
                }
                for (int i = 0; i < output_cells.Count; i++)
                {
                    for (int d = 0; d < reachable_impacts_grid_array[i].Length; d++)
                    {
                        int worksheet_ind = (int)reachable_impacts_grid_array[i][d][0];
                        int row = (int)reachable_impacts_grid_array[i][d][1];
                        int col = (int)reachable_impacts_grid_array[i][d][2];
                        impacts_grid[worksheet_ind][row][col][i] = reachable_impacts_grid_array[i][d][3];
                    }
                }

                //Now we want to average the z-score of every input and store it
                double[][][] average_z_scores = new double[Globals.ThisAddIn.Application.Worksheets.Count][][];
                foreach (Excel.Worksheet worksheet in Globals.ThisAddIn.Application.Worksheets)
                {
                    average_z_scores[worksheet.Index - 1] = new double[worksheet.UsedRange.Rows.Count + worksheet.UsedRange.Row][];
                    for (int row = 0; row < (worksheet.UsedRange.Rows.Count + worksheet.UsedRange.Row); row++)
                    {
                        average_z_scores[worksheet.Index - 1][row] = new double[worksheet.UsedRange.Columns.Count + worksheet.UsedRange.Column];
                    }
                }
                foreach (Excel.Worksheet worksheet in Globals.ThisAddIn.Application.Worksheets)
                {
                    for (int row = 0; row < (worksheet.UsedRange.Rows.Count + worksheet.UsedRange.Row); row++)
                    {
                        for (int col = 0; col < (worksheet.UsedRange.Columns.Count + worksheet.UsedRange.Column); col++)
                        {
                            //If this cell has been perturbed, find it's average z-score
                            double total_z_score = 0.0;
                            double total_output_weight = 0.0;
                            if (times_perturbed[worksheet.Index - 1][row][col] != 0)
                            {
                                for (int i = 0; i < output_cells.Count; i++)
                                {
                                    total_output_weight += output_cells[i].getWeight();
                                    if (impacts_grid[worksheet.Index - 1][row][col][i] != 0)
                                    {
                                        total_z_score += impacts_grid[worksheet.Index - 1][row][col][i] * output_cells[i].getWeight();
                                    }
                                }
                                if (total_output_weight != 0.0)
                                {
                                    average_z_scores[worksheet.Index - 1][row][col] = (total_z_score / total_output_weight);
                                }
                            }
                        }
                    }
                }

                //Look for outliers:
                List<int[]> outliers = new List<int[]>();
                //for (int i = 0; i < output_cells.Count; i++)
                //{
                //    foreach (Excel.Worksheet worksheet in Globals.ThisAddIn.Application.Worksheets)
                //    {
                //        for (int row = 0; row < worksheet.UsedRange.Rows.Count; row++)
                //        {
                //            for (int col = 0; col < worksheet.UsedRange.Columns.Count; col++)
                //            {
                //                if (times_perturbed[worksheet.Index - 1][row][col] != 0)
                //                {
                //                    if (impacts_grid[worksheet.Index - 1][row][col][i] > 2.0)
                //                    {
                //                        //System.Windows.Forms.MessageBox.Show(worksheet.Name + ":R" + (row + 1) + "C" + (col + 1) + " is an outlier with respect to output " + (i + 1) + " with a z-score of " + impacts_grid[worksheet.Index - 1][row][col][i]);
                //                        int[] outlier = new int[3];
                //                        outlier[0] = worksheet.Index - 1;
                //                        outlier[1] = row;
                //                        outlier[2] = col;
                //                        outliers.Add(outlier);
                //                        //worksheet.Cells[row + 1, col + 1].Interior.Color = System.Drawing.Color.Red;
                //                        //return;
                //                    }
                //                }
                //            }
                //        }
                //    }
                //}
                //string outlier_percentages = "Range size:\tOutlier percentage:\n";
                for (int i = 0; i < output_cells.Count; i++)
                {
                    //int outliers_for_this_output = 0; 
                    for (int d = 0; d < reachable_impacts_grid_array[i].Length; d++)
                    {
                        //input_cells_in_computation_count++;
                        int worksheet_ind = (int)reachable_impacts_grid_array[i][d][0];
                        int row = (int)reachable_impacts_grid_array[i][d][1];
                        int col = (int)reachable_impacts_grid_array[i][d][2];
                        //Standard deviations cutoff: 
                        double standard_deviations_cutoff = 2.0;
                        if (reachable_impacts_grid_array[i][d][3] > standard_deviations_cutoff)
                        {
                            //System.Windows.Forms.MessageBox.Show(worksheet.Name + ":R" + (row + 1) + "C" + (col + 1) + " is an outlier with respect to output " + (i + 1) + " with a z-score of " + impacts_grid[worksheet.Index - 1][row][col][i]);
                            int[] outlier = new int[3];
                            bool already_added = false; 
                            outlier[0] = worksheet_ind;
                            outlier[1] = row;
                            outlier[2] = col;
                            //Prevent double-counting of outliers that have already been flagged. 
                            foreach (int[] o in outliers)
                            {
                                if (o[0] == outlier[0] && o[1] == outlier[1] && o[2] == outlier[2])
                                {
                                    already_added = true; 
                                }
                            }
                            if (!already_added)
                            {
                                outliers.Add(outlier);
                            }
                        }
                    }
                }

                //Find the highest weighted average z-score among the outliers
                double max_weighted_z_score = 0.0;
                int[][] outliers_array = outliers.ToArray();
                outliers_count = outliers_array.Length;
                for (int i = 0; i < outliers_array.Length; i++)
                {
                    if (average_z_scores[outliers_array[i][0]][outliers_array[i][1]][outliers_array[i][2]] > max_weighted_z_score)
                    {
                        max_weighted_z_score = average_z_scores[outliers_array[i][0]][outliers_array[i][1]][outliers_array[i][2]];
                    }
                }

                //Color the outliers:
                for (int i = 0; i < outliers_array.Length; i++)
                {
                    Excel.Worksheet worksheet = null;
                    int row = outliers_array[i][1];
                    int col = outliers_array[i][2];
                    //Find the worksheet where this outlier resides
                    foreach (Excel.Worksheet ws in Globals.ThisAddIn.Application.Worksheets)
                    {
                        if (ws.Index - 1 == outliers_array[i][0])
                        {
                            worksheet = ws;
                            break;
                        }
                    }
                    worksheet.Cells[row + 1, col + 1].Interior.Color = System.Drawing.Color.FromArgb(Convert.ToInt32(255 - (average_z_scores[worksheet.Index - 1][row][col] / max_weighted_z_score) * 255), 255, 255);
                }

                impact_scoring_timespan = global_stopwatch.Elapsed;
            }

            Globals.ThisAddIn.Application.ScreenUpdating = true;
            
            // Format and display the TimeSpan value. 
            string tree_building_time = tree_building_timespan.TotalSeconds + ""; //String.Format("{0:00}:{1:00}.{2:00}", tree_building_timespan.Minutes, tree_building_timespan.Seconds, tree_building_timespan.Milliseconds / 10);
            string swapping_time = (swapping_timespan.TotalSeconds - tree_building_timespan.TotalSeconds) + ""; //String.Format("{0:00}:{1:00}.{2:00}", swapping_timespan.Minutes, swapping_timespan.Seconds, swapping_timespan.Milliseconds / 10);
            string impact_scoring_time = (impact_scoring_timespan.TotalSeconds - swapping_timespan.TotalSeconds) + ""; //String.Format("{0:00}:{1:00}.{2:00}", z_score_timespan.Minutes, z_score_timespan.Seconds, z_score_timespan.Milliseconds / 10);
            //string average_z_score_time = average_z_score_timespan.TotalSeconds + ""; //String.Format("{0:00}:{1:00}.{2:00}", average_z_score_timespan.Minutes, average_z_score_timespan.Seconds, average_z_score_timespan.Milliseconds / 10);
            //string outlier_detection_time = outlier_detection_timespan.TotalSeconds + ""; //String.Format("{0:00}:{1:00}.{2:00}", outlier_detection_timespan.Minutes, outlier_detection_timespan.Seconds, outlier_detection_timespan.Milliseconds / 10);
            //string outlier_coloring_time = outlier_coloring_timespan.TotalSeconds + ""; //String.Format("{0:00}:{1:00}.{2:00}", outlier_coloring_timespan.Minutes, outlier_coloring_timespan.Seconds, outlier_coloring_timespan.Milliseconds / 10);
            //System.Windows.Forms.MessageBox.Show("Done building dependence graph.\nTime elapsed: " + treeTime);
            global_stopwatch.Stop();
            // Get the elapsed time as a TimeSpan value.
            TimeSpan global_timespan = global_stopwatch.Elapsed;
            //string global_time = String.Format("{0:00}:{1:00}:{2:00}.{3:00}", global_timespan.Hours, global_timespan.Minutes, global_timespan.Seconds, global_timespan.Milliseconds / 10);
            string global_time = global_timespan.TotalSeconds + ""; //(tree_building_timespan.TotalSeconds + swapping_timespan.TotalSeconds + z_score_timespan.TotalSeconds + average_z_score_timespan.TotalSeconds + outlier_detection_timespan.TotalSeconds + outlier_coloring_timespan.TotalSeconds) + ""; //String.Format("{0:00}:{1:00}.{2:00}",
            //    tree_building_timespan.Minutes + swapping_timespan.Minutes + z_score_timespan.Minutes + average_z_score_timespan.Minutes + outlier_detection_timespan.Minutes + outlier_coloring_timespan.Minutes,
            //    tree_building_timespan.Seconds + swapping_timespan.Seconds + z_score_timespan.Seconds + average_z_score_timespan.Seconds + outlier_detection_timespan.Seconds + outlier_coloring_timespan.Seconds,  
            //    (tree_building_timespan.Milliseconds + swapping_timespan.Milliseconds + z_score_timespan.Milliseconds + average_z_score_timespan.Milliseconds + outlier_detection_timespan.Milliseconds + outlier_coloring_timespan.Milliseconds) / 10);
            
            Display timeDisplay = new Display();
            stats_text += "" //+ "Benchmark:\tNumber of formulas:\tRaw input count:\tInputs to computations:\tTotal (s):\tTree Construction (s):\tSwapping (s):\tZ-Score Calculation (s):\t"
//          //  + "Outlier Detection (s):\tOutlier Coloring (s):\t"
            //+ "Outliers found:\n"
//                //"Formula cells:\t" + formula_cells_count + "\n"
//                //+ "Number of input cells involved in computations:\t" + input_cells_in_computation_count
//                //+ "\nExecution times (seconds): "
                + Globals.ThisAddIn.Application.ActiveWorkbook.Name + "\t"
                + formula_cells_count + "\t"
                + raw_input_cells_in_computation_count + "\t"
                + input_cells_in_computation_count + "\t"
                + global_time + "\t"
                + tree_building_time + "\t"
                + swapping_time + "\t"
                + impact_scoring_time + "\t"
                + outliers_count;
//                //+ (z_score_time + average_z_score_time) + "\t"
//                //+ outlier_detection_time + "\t"
//                //+ outlier_coloring_time;
////                + "\nTotal execution time: " + global_time + ":\n"
////                + "\tTree construction time: " + tree_building_time + "\n"
////                + "\tSwapping time: " + swapping_time + "\n"
////                + "\tZ-score calculation time: " + z_score_time + "\n"
////                + "\tWeighted average z-score calculation time: " + average_z_score_time + "\n"
////                + "\tOutlier detection time: " + outlier_detection_time + "\n"
////                + "\tOutlier coloring time: " + outlier_coloring_time + "\n";
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


            //Print out text for GraphViz representation of the dependence graph
            //string tree = "";
            //string ranges_text = "";
            //foreach (TreeNode[][] node_arr_arr in nodes_grid)
            //{
            //    foreach (TreeNode[] node_arr in node_arr_arr)
            //    {
            //        foreach (TreeNode node in node_arr)
            //        {
            //            if (node != null)
            //            {
            //                tree += node.toGVString(0) + "\n"; //tree += node.toGVString(max_weight) + "\n";
            //            }
            //        }
            //    }
            //}
            //foreach (TreeNode node in ranges)
            //{
            //    tree += node.toGVString(0) + "\n"; //tree += node.toGVString(max_weight) + "\n";
            //    foreach (TreeNode parent in node.getParents())
            //    {
            //        ranges_text += parent.getWorksheetObject().Index + "," + parent.getName().Replace("$", "") + "," + parent.getWorksheetObject().get_Range(parent.getName()).Value + "\n";
            //    }
            //}
            
            //Display disp = new Display();
            //disp.textBox1.Text = "digraph g{" + tree + "}";
            //disp.ShowDialog();
            //Display disp_ranges = new Display();
            //disp_ranges.textBox1.Text = ranges_text;
            //disp_ranges.ShowDialog();
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

        private void checkBox1_Click(object sender, RibbonControlEventArgs e)
        {
        }


        //Button for testing random code :)
        private void button7_Click(object sender, RibbonControlEventArgs e)
        {
            System.Windows.Forms.MessageBox.Show(Globals.ThisAddIn.Application.ActiveWorkbook.Path + "");
            System.Windows.Forms.MessageBox.Show(Globals.ThisAddIn.Application.Workbooks[1] + "");
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
    }
}
