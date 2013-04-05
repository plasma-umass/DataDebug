using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
using System.Text.RegularExpressions;
using TreeDict = System.Collections.Generic.Dictionary<AST.Address, DataDebugMethods.TreeNode>;
using TreeList = System.Collections.Generic.List<DataDebugMethods.TreeNode>;
using TreeDictPair = System.Collections.Generic.KeyValuePair<AST.Address, DataDebugMethods.TreeNode>;
using Microsoft.FSharp.Core;

namespace DataDebugMethods
{
    public static class ConstructTree
    {
        /*
         * This method constructs the dependency graph from the worksheet.
         * It analyzes formulas and looks for references to cells or ranges of cells.
         * It also looks for any charts, and adds those to the dependency graph as well. 
         * After the dependency graph is constructed, we use it to determine and propagate weights to all nodes in the graph. 
         * This method also contains the perturbation procedure and outlier analysis logic.
         * In the end, a text representation of the dependency graph is given in GraphViz format. It includes the entire graph and the weights of the nodes.
         */
        public static void constructTree(AnalysisData analysisData, Excel.Application app)
        {
            Excel.Sheets ws = app.Worksheets;

            analysisData.pb.SetProgress(0);
            analysisData.input_cells_in_computation_count = 0;
            analysisData.raw_input_cells_in_computation_count = 0;

            // Get a range representing the formula cells for each worksheet in each workbook
            ArrayList formulaRanges = ConstructTree.GetFormulaRanges(ws, app);
            analysisData.formula_cells_count = ConstructTree.CountFormulaCells(formulaRanges);

            // Create nodes for every cell containing a formula
            analysisData.formula_nodes = ConstructTree.CreateFormulaNodes(formulaRanges, app);

            //Now we parse the formulas in nodes to extract any range and cell references
            foreach(TreeDictPair pair in analysisData.formula_nodes)
            {
                // This is a formula:
                TreeNode formula_node = pair.Value;

                // For each of the ranges found in the formula by the parser,
                // 1. make a new TreeNode for the range
                // 2. make TreeNodes for each of the cells in that range
                foreach (Excel.Range input_range in ExcelParserUtility.GetReferencesFromFormula(formula_node.getFormula(), formula_node.getWorkbookObject(), formula_node.getWorksheetObject()))
                {
                    // this function both creates a TreeNode and adds it to AnalysisData.input_ranges
                    TreeNode range_node = ConstructTree.MakeRangeTreeNode(analysisData.input_ranges, input_range, formula_node);
                    // this function both creates cell TreeNodes for a range and adds it to AnalysisData.cell_nodes
                    ConstructTree.CreateCellNodesFromRange(range_node, formula_node, analysisData.formula_nodes, analysisData.cell_nodes);
                }

                // For each single-cell input found in the formula by the parser,
                // link to output TreeNode if the input cell is a formula. This allows
                // us to consider functions with single-cell inputs as outputs.
                foreach (AST.Address input_addr in ExcelParserUtility.GetSingleCellReferencesFromFormula(formula_node.getFormula(), formula_node.getWorkbookObject(), formula_node.getWorksheetObject()))
                {
                    // Find the input cell's TreeNode; if there isn't one, move on.
                    // We don't care about scalar inputs that are not functions
                    TreeNode tn;
                    if (analysisData.formula_nodes.TryGetValue(input_addr, out tn))
                    {
                        // sanity check-- should be a formula
                        if (tn.isFormula())
                        {
                            // link input to output formula node
                            tn.addChild(formula_node);
                            formula_node.addParent(tn);
                        }
                    }
                }
            }

            //TODO -- we are not able to capture ranges that are identified in stored procedures or macros, just ones referenced in formulas
            //TODO -- Dealing with fuzzing of charts -- idea: any cell that feeds into a chart is essentially an output; the chart is just a visual representation (can charts operate on values before they are displayed? don't think so...)
            ConstructTree.StoreOutputs(analysisData);
        }

        public static int CountFormulaCells(ArrayList rs)
        {
            int count = 0;
            foreach (Excel.Range r in rs)
            {
                if (r != null)
                {
                    count += r.Cells.Count;
                }
            }
            return count;
        }

        public static ArrayList GetFormulaRanges(Excel.Sheets ws, Excel.Application app)
        {
            // This keeps track of the range to be analyzed in every worksheet of the workbook
            // We have to use ArrayList because COM interop does not work with generics.
            var analysisRanges = new ArrayList(); 

            foreach (Excel.Worksheet w in ws)
            {
                Excel.Range formula_cells = null;
                // iterate over all of the cells in a particular worksheet
                // these actually are cells, because that's what you get when you
                // iterate over the UsedRange property
                foreach (Excel.Range cell in w.UsedRange)
                {
                    // the cell thinks it has a formula
                    if (cell.HasFormula)
                    {
                        // this is our first time around; formula_cells is not yet set,
                        // so set it
                        if (formula_cells == null)
                        {
                            formula_cells = cell;
                        }
                        // it's not our first time around, so union the current cell with
                        // the previously found formula cell
                        else
                        {
                            formula_cells = app.Union(cell, formula_cells);
                        }
                    }
                }
                // we found at least one cell
                if (formula_cells != null)
                {
                    analysisRanges.Add(formula_cells);
                }
            }
            return analysisRanges;
        }

        //First we create nodes for every non-null cell; then we will operate on these node objects, connecting them in the tree, etc. 
        //This includes cells that contain constants and formulas
        //Go through every worksheet
        public static TreeDict CreateFormulaNodes(ArrayList rs, Excel.Application app)
        {
            Excel.Workbook wb = app.ActiveWorkbook;

            // init nodes
            var nodes = new TreeDict();

            foreach (Excel.Range worksheet_range in rs)
            {
                foreach (Excel.Range cell in worksheet_range)
                {
                    if (cell.Value2 != null)
                    {
                        var addr = ExcelParser.GetAddress(cell.Address[true, true, Excel.XlReferenceStyle.xlR1C1, false], wb, cell.Worksheet);
                        var n = new TreeNode(cell, cell.Address, cell.Worksheet, wb);
                        
                        if (cell.HasFormula)
                        {
                            n.setIsFormula();
                            n.DontPerturb();
                            if (cell.Formula == null)
                            {
                                throw new Exception("null formula!!! argh!!!");
                            }
                            n.setFormula(cell.Formula);
                            nodes.Add(addr, n);
                        }
                    }
                }
            }
            return nodes;
        }

        public static void StoreOutputs(AnalysisData analysisData)
        {
            // Collect output values
            foreach (TreeDictPair tdp in analysisData.formula_nodes)
            {
                var node = tdp.Value;
                if (!node.hasChildren() && node.hasParents()) //Nodes that do not feed into any other nodes are considered output, unless nothing feeds into them either. 
                {
                    analysisData.output_cells.Add(node);
                }
            }

            //This part stores all the output values before any perturbations are applied
            foreach (TreeNode n in analysisData.output_cells)
            {
                // If the TreeNode is a chart
                if (n.isChart())
                {
                    // Add a StartValue with the average of the range of inputs for each range of inputs
                    double sum = 0.0;
                    TreeNode parent_range = n.getParents()[0];
                    foreach (TreeNode par in parent_range.getParents())
                    {
                        sum = sum + par.getWorksheetObject().get_Range(par.getName()).Value;
                    }
                    double average = sum / parent_range.getParents().Count;
                    StartValue sv = new StartValue(average);
                    analysisData.starting_outputs.Add(sv);

                }
                // If the TreeNode is a cell
                else
                {
                    Excel.Worksheet nodeWorksheet = n.getWorksheetObject(); //This is be the worksheet where the node n is located
                    Excel.Range cell = nodeWorksheet.get_Range(n.getName());
                    try     //If the output is a number
                    {
                        double d = (double)nodeWorksheet.get_Range(n.getName()).Value;
                        StartValue sv = new StartValue(d);
                        analysisData.starting_outputs.Add(sv); //Try adding it as a number
                    }
                    catch   //If the output is a string
                    {
                        string s = (string)nodeWorksheet.get_Range(n.getName()).Value.ToString();
                        StartValue sv = new StartValue(s);
                        analysisData.starting_outputs.Add(sv); //If it doesn't work, it must be a string output
                    }
                }
            }
        }

        public static string GenerateGraphVizTree(TreeDict nodes)
        {
            string tree = "";
            foreach (TreeDictPair nodePair in nodes)
            {
                tree += nodePair.Value.toGVString(0.0) + "\n";
            }
            return "digraph g{" + tree + "}"; 
        }

        public static void setUpGrids(AnalysisData analysisData)
        {
            analysisData.influences_grid = new double[analysisData.worksheets.Count + analysisData.charts.Count][][];
            analysisData.times_perturbed = new int[analysisData.worksheets.Count + analysisData.charts.Count][][];
            foreach (Excel.Worksheet worksheet in analysisData.worksheets)
            {
                analysisData.influences_grid[worksheet.Index - 1] = new double[worksheet.UsedRange.Rows.Count + worksheet.UsedRange.Row][];
                analysisData.times_perturbed[worksheet.Index - 1] = new int[worksheet.UsedRange.Rows.Count + worksheet.UsedRange.Row][];
                for (int row = 0; row < (worksheet.UsedRange.Rows.Count + worksheet.UsedRange.Row); row++)
                {
                    analysisData.influences_grid[worksheet.Index - 1][row] = new double[worksheet.UsedRange.Columns.Count + worksheet.UsedRange.Column];
                    analysisData.times_perturbed[worksheet.Index - 1][row] = new int[worksheet.UsedRange.Columns.Count + worksheet.UsedRange.Column];
                    for (int col = 0; col < (worksheet.UsedRange.Columns.Count + worksheet.UsedRange.Column); col++)
                    {
                        analysisData.influences_grid[worksheet.Index - 1][row][col] = 0.0;
                        analysisData.times_perturbed[worksheet.Index - 1][row][col] = 0;
                    }
                }
            }
        }

        public static void SwappingProcedure(AnalysisData analysisData)
        {
            foreach (TreeNode range_node in analysisData.input_ranges)
            {
                //For every range node
                double[] influences = new double[range_node.getParents().Count]; //Array to keep track of the influence values for every cell in the range
                
                int swaps_per_range = 30;
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
                    analysisData.input_cells_in_computation_count++;

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
                            analysisData.times_perturbed[cell.Worksheet.Index - 1][cell.Row - 1][cell.Column - 1]++;
                        }
                        catch
                        {
                            cell.Interior.Color = System.Drawing.Color.Purple;
                        }
                        Excel.Range sibling_cell = sibling.getWorksheetObject().get_Range(sibling.getName());
                        cell.Value = sibling_cell.Value; //This is the swap -- we assign the value of the sibling cell to the value of our cell
                        delta = 0.0;
                        //foreach (TreeNode n in output_cells)
                        for (int i = 0; i < analysisData.output_cells.Count; i++)
                        {
                            try
                            {
                                //If this output is not reachable from this cell, continue
                                if (analysisData.reachable_grid[cell.Worksheet.Index - 1][cell.Row - 1][cell.Column - 1][i] == false)
                                {
                                    continue;
                                }
                            }
                            catch
                            {
                                continue;
                            }
                            TreeNode n = analysisData.output_cells[i];
                            if (analysisData.starting_outputs[i].get_string() == null) // If the output is not a string
                            {
                                if (!n.isChart())   //If the output is not a chart, it must be a number
                                {
                                    delta = Math.Abs(analysisData.starting_outputs[i].get_double() - (double)n.getWorksheetObject().get_Range(n.getName()).Value);  //Compute the absolute change caused by the swap
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
                                    delta = Math.Abs(analysisData.starting_outputs[i].get_double() - average);
                                }
                            }
                            else  // If the output is a string
                            {
                                if (String.Equals(analysisData.starting_outputs[i].get_string(), n.getWorksheetObject().get_Range(n.getName()).Value, StringComparison.Ordinal))
                                {
                                    delta = 0.0;
                                }
                                else
                                {
                                    delta = 1.0;
                                }
                            }
                            //Add to the impact of the cell for this output
                            analysisData.impacts_grid[cell.Worksheet.Index - 1][cell.Row - 1][cell.Column - 1][i] += delta;
                            //Compare the min/max values for this output to this delta
                            if (analysisData.min_max_delta_outputs[i][0] == -1.0)
                            {
                                analysisData.min_max_delta_outputs[i][0] = delta;
                            }
                            else
                            {
                                if (analysisData.min_max_delta_outputs[i][0] > delta)
                                {
                                    analysisData.min_max_delta_outputs[i][0] = delta;
                                }
                            }
                            if (analysisData.min_max_delta_outputs[i][1] < delta)
                            {
                                analysisData.min_max_delta_outputs[i][1] = delta;
                            }
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

            //Populate reachable_impacts_grid_array from impacts_grid
            for (int i = 0; i < analysisData.output_cells.Count; i++)
            {
                for (int d = 0; d < analysisData.reachable_impacts_grid_array[i].Length; d++)
                {
                    analysisData.reachable_impacts_grid_array[i][d] = new double[4] { analysisData.reachable_impacts_grid_array[i][d][0], 
                            analysisData.reachable_impacts_grid_array[i][d][1], 
                            analysisData.reachable_impacts_grid_array[i][d][2], 
                            analysisData.impacts_grid[(int)analysisData.reachable_impacts_grid_array[i][d][0]][(int)analysisData.reachable_impacts_grid_array[i][d][1]][(int)analysisData.reachable_impacts_grid_array[i][d][2]][i] };
                }
            }
            //System.Windows.Forms.MessageBox.Show("Input cells in computation: " + analysisData.input_cells_in_computation_count);
        }

        public static void ComputeZScoresAndFindOutliers(AnalysisData analysisData)
        {
            //Now for each output, compute the z-score of the impact of each input
            for (int i = 0; i < analysisData.output_cells.Count; i++)
            {
                //Find the mean for the output
                double output_sum = 0.0;

                for (int d = 0; d < analysisData.reachable_impacts_grid_array[i].Length; d++)
                {
                    int worksheet_ind = (int)analysisData.reachable_impacts_grid_array[i][d][0];
                    int row = (int)analysisData.reachable_impacts_grid_array[i][d][1];
                    int col = (int)analysisData.reachable_impacts_grid_array[i][d][2];
                    if (analysisData.times_perturbed[worksheet_ind][row][col] != 0)
                    {
                        output_sum += analysisData.impacts_grid[worksheet_ind][row][col][i];
                    }
                }

                double output_average = 0.0;
                if (analysisData.reachable_impacts_grid_array[i].Length != 0)
                {
                    output_average = output_sum / (double)analysisData.reachable_impacts_grid_array[i].Length;
                }
                else  //if none of the entries can reach this output, all impacts must be equal to 0.
                {
                    output_average = 0.0;
                }
                //Find the sample standard deviation for this output
                double variance = 0.0;

                for (int d = 0; d < analysisData.reachable_impacts_grid_array[i].Length; d++)
                {
                    int worksheet_ind = (int)analysisData.reachable_impacts_grid_array[i][d][0];
                    int row = (int)analysisData.reachable_impacts_grid_array[i][d][1];
                    int col = (int)analysisData.reachable_impacts_grid_array[i][d][2];
                    if (analysisData.times_perturbed[worksheet_ind][row][col] != 0)
                    {
                        variance += Math.Pow(output_average - analysisData.impacts_grid[worksheet_ind][row][col][i], 2) / (double)analysisData.reachable_impacts_grid_array[i].Length;
                    }
                }
                double std_dev = Math.Sqrt(variance);

                for (int d = 0; d < analysisData.reachable_impacts_grid_array[i].Length; d++)
                {
                    int worksheet_ind = (int)analysisData.reachable_impacts_grid_array[i][d][0];
                    int row = (int)analysisData.reachable_impacts_grid_array[i][d][1];
                    int col = (int)analysisData.reachable_impacts_grid_array[i][d][2];
                    if (analysisData.times_perturbed[worksheet_ind][row][col] != 0)
                    {
                        if (std_dev != 0.0)
                        {
                            analysisData.reachable_impacts_grid_array[i][d][3] = Math.Abs((analysisData.impacts_grid[worksheet_ind][row][col][i] - output_average) / std_dev);
                        }
                        else  //if std_dev == 0.0
                        {
                            //If the standard deviation is zero, then all the impacts were the same and we shouldn't flag any entries, so set their z-scores to 0.0
                            analysisData.reachable_impacts_grid_array[i][d][3] = 0.0;
                        }
                    }
                }
            }

            //Repopulate impacts_grid with the z-scores from reachable_impacts_grid_array
            foreach (Excel.Worksheet worksheet in analysisData.worksheets)
            {
                Excel.Range used_range = worksheet.get_Range("A1");
                for (int row = 0; row < (worksheet.UsedRange.Rows.Count + worksheet.UsedRange.Row); row++)
                {
                    for (int col = 0; col < (worksheet.UsedRange.Columns.Count + worksheet.UsedRange.Column); col++)
                    {
                        for (int i = 0; i < analysisData.output_cells.Count; i++)
                        {
                            analysisData.impacts_grid[worksheet.Index - 1][row][col][i] = 0.0;
                        }
                    }
                }
            }
            for (int i = 0; i < analysisData.output_cells.Count; i++)
            {
                for (int d = 0; d < analysisData.reachable_impacts_grid_array[i].Length; d++)
                {
                    int worksheet_ind = (int)analysisData.reachable_impacts_grid_array[i][d][0];
                    int row = (int)analysisData.reachable_impacts_grid_array[i][d][1];
                    int col = (int)analysisData.reachable_impacts_grid_array[i][d][2];
                    analysisData.impacts_grid[worksheet_ind][row][col][i] = analysisData.reachable_impacts_grid_array[i][d][3];
                }
            }

            //Now we want to average the z-score of every input and store it
            double[][][] average_z_scores = new double[analysisData.worksheets.Count][][];
            foreach (Excel.Worksheet worksheet in analysisData.worksheets)
            {
                average_z_scores[worksheet.Index - 1] = new double[worksheet.UsedRange.Rows.Count + worksheet.UsedRange.Row][];
                for (int row = 0; row < (worksheet.UsedRange.Rows.Count + worksheet.UsedRange.Row); row++)
                {
                    average_z_scores[worksheet.Index - 1][row] = new double[worksheet.UsedRange.Columns.Count + worksheet.UsedRange.Column];
                }
            }
            foreach (Excel.Worksheet worksheet in analysisData.worksheets)
            {
                for (int row = 0; row < (worksheet.UsedRange.Rows.Count + worksheet.UsedRange.Row); row++)
                {
                    for (int col = 0; col < (worksheet.UsedRange.Columns.Count + worksheet.UsedRange.Column); col++)
                    {
                        //If this cell has been perturbed, find it's average z-score
                        double total_z_score = 0.0;
                        double total_output_weight = 0.0;
                        if (analysisData.times_perturbed[worksheet.Index - 1][row][col] != 0)
                        {
                            for (int i = 0; i < analysisData.output_cells.Count; i++)
                            {
                                total_output_weight += analysisData.output_cells[i].getWeight();
                                if (analysisData.impacts_grid[worksheet.Index - 1][row][col][i] != 0)
                                {
                                    total_z_score += analysisData.impacts_grid[worksheet.Index - 1][row][col][i] * analysisData.output_cells[i].getWeight();
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
            System.Collections.Generic.List<int[]> outliers = new System.Collections.Generic.List<int[]>();
            for (int i = 0; i < analysisData.output_cells.Count; i++)
            {
                //int outliers_for_this_output = 0; 
                for (int d = 0; d < analysisData.reachable_impacts_grid_array[i].Length; d++)
                {
                    //input_cells_in_computation_count++;
                    int worksheet_ind = (int)analysisData.reachable_impacts_grid_array[i][d][0];
                    int row = (int)analysisData.reachable_impacts_grid_array[i][d][1];
                    int col = (int)analysisData.reachable_impacts_grid_array[i][d][2];
                    //Standard deviations cutoff: 
                    double standard_deviations_cutoff = 2.0;
                    if (analysisData.reachable_impacts_grid_array[i][d][3] > standard_deviations_cutoff)
                    {
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
            analysisData.outliers_count = outliers_array.Length;
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
                foreach (Excel.Worksheet ws in analysisData.worksheets)
                {
                    if (ws.Index - 1 == outliers_array[i][0])
                    {
                        worksheet = ws;
                        break;
                    }
                }
                worksheet.Cells[row + 1, col + 1].Interior.Color = System.Drawing.Color.FromArgb(Convert.ToInt32(255 - (average_z_scores[worksheet.Index - 1][row][col] / max_weighted_z_score) * 255), 255, 255);
                analysisData.oldToolOutlierAddresses.Add(worksheet.Cells[row + 1, col + 1].Address);
            }

        }

        public static void CreateCellNodesFromRange(TreeNode rangeNode, TreeNode formulaNode, TreeDict formula_nodes, TreeDict cell_nodes)
        {
            foreach (Excel.Range cell in rangeNode.getCOMObject())
            {
                TreeNode cellNode = null;
                //See if there is an existing node for this cell already in nodes; if there is, do not add it again - just grab the existing one
                if (!formula_nodes.TryGetValue(ExcelParser.GetAddress(cell.Address[true, true, Excel.XlReferenceStyle.xlR1C1, false], formulaNode.getWorkbookObject(), cell.Worksheet), out cellNode))
                {
                    //TODO CORRECT THE WORKBOOK PARAMETER IN THIS LINE: (IT SHOULD BE THE WORKBOOK OF cell, WHICH SHOULD COME FROM GetReferencesFromFormula
                    var addr = ExcelParser.GetAddress(cell.Address[true, true, Excel.XlReferenceStyle.xlR1C1, false], formulaNode.getWorkbookObject(), cell.Worksheet);
                    cellNode = new TreeNode(cell, cell.Address, cell.Worksheet, formulaNode.getWorkbookObject());

                    // Add to cell_nodes, not formula_nodes
                    if (!cell_nodes.ContainsKey(addr))
                    {
                        cell_nodes.Add(addr, cellNode);
                    }
                    else
                    {
                        cellNode = cell_nodes[addr];
                    }
                }

                // check whether this cell is a formula; if it is, set the "don't perturb" bit on the range
                if (cell.HasFormula)
                {
                    // TODO: check for constant formulas... if that's the case, then we again consider the input to be a formula
                    rangeNode.DontPerturb();
                }
                // register this cell as an input to the rangeNode
                rangeNode.addParent(cellNode);
                // register the range as an output of this input cell
                cellNode.addChild(formulaNode);
                // register this cell as an input to the formulaNode
                formulaNode.addParent(cellNode);
            }
        } // CreateCellNodesFromRange ends here

        public static TreeNode MakeRangeTreeNode(TreeList ranges, Excel.Range range, TreeNode node)
        {
            TreeNode rangeNode = null;
            // See if there is an existing node for this range already in "ranges";
            // if there is, do not add it again - just grab the existing one
            foreach (TreeNode existingNode in ranges)
            {
                if (existingNode.getName().Equals(range.Address))
                {
                    rangeNode = existingNode;
                    break;
                }
            }
            if (rangeNode == null)
            {
                rangeNode = new TreeNode(range, range.Address, range.Worksheet, node.getWorkbookObject());
                ranges.Add(rangeNode);
            }
            return rangeNode;
        }

        public static bool AllValuesBelowLength(int len, TreeDict ts)
        {
            foreach (KeyValuePair<AST.Address, TreeNode> pair in ts)
            {
                if (pair.Value.getCOMValueAsString() != null &&
                    pair.Value.getCOMValueAsString().Length > len)
                {
                     return false;
                }
            }
            return true;
        }

        public static FSharpOption<TurkJob[]> DataForMTurk(Excel.Application app, int maxlen)
        {
            const int WIDTH = 10;

            AnalysisData data = new AnalysisData(app);
            data.Reset();
            ConstructTree.constructTree(data, app);
            data.pb.Close();

            // sanity check
            if (!AllValuesBelowLength(maxlen, data.cell_nodes))
            {
                return FSharpOption<TurkJob[]>.None;
            }

            // determine # of rows
            int rows;
            if (data.cell_nodes.Count % WIDTH > 0)
            {
                rows = data.cell_nodes.Count / WIDTH + 1;
            }
            else
            {
                rows = data.cell_nodes.Count / WIDTH;
            }
            var output = new string[rows, WIDTH];
            var addrs = new string[rows, WIDTH];

            // split data
            var j = 0;
            var i = 0;
            foreach (KeyValuePair<AST.Address,TreeNode> pair in data.cell_nodes)
            {
                var t = pair.Value;
                if (t.getCOMValueAsString() != null)
                {
                    output[i,j] = t.getCOMValueAsString();
                }
                else
                {
                    output[i,j] = "";
                }
                addrs[i,j] = t.getCOMObject().Address;
                j = (j + 1) % WIDTH;
                if (j == 0)
                {
                    i++;
                }
            }

            // pad with empties, if necessary
            while (j != 0)
            {
                output[i,j] = "ABRAHAMLINCOLN";
                addrs[i,j] = "ZAA221";
                j = (j + 1) % WIDTH;
            }

            return FSharpOption<TurkJob[]>.Some(ToMTurkJob(output, addrs));
        }

        public static string Truncate(string str, int len)
        {
            if (str.Length <= len)
            {
                return str;
            }
            else
            {
                return str.Substring(0, len);
            }
        }

        public static TurkJob[] ToMTurkJob(string[,] data, string[,] addrs)
        {
            var jobs = new TurkJob[data.GetLength(0)];

            for (var job_id = 0; job_id < data.GetLength(0); job_id++)
            {
                var tj = new TurkJob();
                tj.SetJobId(job_id);
                var tjdata = new string[data.GetLength(1)];
                var tjaddrs = new string[addrs.GetLength(1)];
                for(var i = 0; i < data.GetLength(1); i++)
                {
                    tjdata[i] = data[job_id, i];
                    tjaddrs[i] = addrs[job_id, i];
                }
                tj.SetCells(tjdata);
                tj.SetAddrs(tjaddrs);
                jobs[job_id] = tj;
            }

            return jobs;
        }

    } // ConstructTree class ends here
} // namespace ends here
