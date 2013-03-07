using System;
using System.Collections;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
using System.Text.RegularExpressions;
using TreeDict = System.Collections.Generic.Dictionary<AST.Address, DataDebugMethods.TreeNode>;
using TreeList = System.Collections.Generic.List<DataDebugMethods.TreeNode>;
using TreeDictPair = System.Collections.Generic.KeyValuePair<AST.Address, DataDebugMethods.TreeNode>;

namespace DataDebugMethods
{
    public static class ConstructTree
    {
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
                            formula_cells = app.Union(
                                            cell,
                                            formula_cells,
                                            Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                            Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                            Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                            Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
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
                    if (cell.Value != null)
                    {
                        var addr = ExcelParser.GetAddress(cell.Address[true, true, Excel.XlReferenceStyle.xlR1C1, false], wb, cell.Worksheet);
                        var n = new TreeNode(cell.Address, cell.Worksheet, wb);
                        
                        if (cell.HasFormula)
                        {
                            n.setIsFormula();
                            n.setFormula(cell.Formula);
                            nodes.Add(addr, n);
                        }
                    }
                }
            }
            return nodes;
        }

        public static void StripLookups(string formula)
        {
            Regex hlookup_regex = new Regex(@"(HLOOKUP\([A-Za-z0-9_ :\$\f\n\r\t\v]+,[A-Za-z0-9_ :\$\f\n\r\t\v]+,[A-Za-z0-9_ :\$\f\n\r\t\v]+\))", RegexOptions.Compiled);
            Regex hlookup_regex_1 = new Regex(@"(HLOOKUP\([A-Za-z0-9_ :\$\f\n\r\t\v]+,[A-Za-z0-9_ :\$\f\n\r\t\v]+,[A-Za-z0-9_ :\$\f\n\r\t\v]+,[A-Za-z0-9_ :\$\f\n\r\t\v]+\))", RegexOptions.Compiled);
            Regex vlookup_regex = new Regex(@"(VLOOKUP\([A-Za-z0-9_ :\$\f\n\r\t\v]+,[A-Za-z0-9_ :\$\f\n\r\t\v]+,[A-Za-z0-9_ :\$\f\n\r\t\v]+\))", RegexOptions.Compiled);
            Regex vlookup_regex_1 = new Regex(@"(VLOOKUP\([A-Za-z0-9_ :\$\f\n\r\t\v]+,[A-Za-z0-9_ :\$\f\n\r\t\v]+,[A-Za-z0-9_ :\$\f\n\r\t\v]+,[A-Za-z0-9_ :\$\f\n\r\t\v]+\))", RegexOptions.Compiled);
            MatchCollection matchedLookups = hlookup_regex.Matches(formula);
            foreach (Match match in matchedLookups)
            {
                formula = formula.Replace(match.Value, "");
            }
            matchedLookups = hlookup_regex_1.Matches(formula);
            foreach (Match match in matchedLookups)
            {
                formula = formula.Replace(match.Value, "");
            }
            matchedLookups = vlookup_regex.Matches(formula);
            foreach (Match match in matchedLookups)
            {
                formula = formula.Replace(match.Value, "");
            }
            matchedLookups = vlookup_regex_1.Matches(formula);
            foreach (Match match in matchedLookups)
            {
                formula = formula.Replace(match.Value, "");
            }
        }

        public static void StoreOutputs(System.Collections.Generic.List<StartValue> starting_outputs, TreeList output_cells, TreeDict nodes)
        {
            // Collect output values
            foreach (TreeDictPair tdp in nodes)
            {
                var node = tdp.Value;
                if (!node.hasChildren() && node.hasParents()) //Nodes that do not feed into any other nodes are considered output, unless nothing feeds into them either. 
                {
                    output_cells.Add(node);
                }
            }

            //This part stores all the output values before any perturbations are applied
            foreach (TreeNode n in output_cells)
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
                    starting_outputs.Add(sv);

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
                        starting_outputs.Add(sv); //Try adding it as a number
                    }
                    catch   //If the output is a string
                    {
                        string s = (string)nodeWorksheet.get_Range(n.getName()).Value.ToString();
                        StartValue sv = new StartValue(s);
                        starting_outputs.Add(sv); //If it doesn't work, it must be a string output
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

        public static void setUpGrids(ref double[][][] influences_grid, ref int[][][] times_perturbed, Excel.Sheets worksheets, Excel.Sheets charts)
        {
            influences_grid = new double[worksheets.Count + charts.Count][][];
            times_perturbed = new int[worksheets.Count + charts.Count][][];
            foreach (Excel.Worksheet worksheet in worksheets)
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
        }

        public static void SwappingProcedure(TreeList swap_domain, ref int input_cells_in_computation_count, ref double[][] min_max_delta_outputs, ref double[][][][] impacts_grid, ref int[][][] times_perturbed, ref TreeList output_cells, ref bool[][][][] reachable_grid, ref System.Collections.Generic.List<StartValue> starting_outputs, ref double[][][] reachable_impacts_grid_array)
        {
            foreach (TreeNode range_node in swap_domain)
            {
                //If this node is not a range, continue to the next node because no perturbations can be done on this node.
                //if (!range_node.isRange())
                //{
                //    continue;
                //}
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

        }

        public static void ComputeZScoresAndFindOutliers(TreeList output_cells, double[][][] reachable_impacts_grid_array, double[][][][] impacts_grid, int[][][] times_perturbed, Excel.Sheets worksheets, int outliers_count)
        {
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
            foreach (Excel.Worksheet worksheet in worksheets)
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
            double[][][] average_z_scores = new double[worksheets.Count][][];
            foreach (Excel.Worksheet worksheet in worksheets)
            {
                average_z_scores[worksheet.Index - 1] = new double[worksheet.UsedRange.Rows.Count + worksheet.UsedRange.Row][];
                for (int row = 0; row < (worksheet.UsedRange.Rows.Count + worksheet.UsedRange.Row); row++)
                {
                    average_z_scores[worksheet.Index - 1][row] = new double[worksheet.UsedRange.Columns.Count + worksheet.UsedRange.Column];
                }
            }
            foreach (Excel.Worksheet worksheet in worksheets)
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
            System.Collections.Generic.List<int[]> outliers = new System.Collections.Generic.List<int[]>();
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
                foreach (Excel.Worksheet ws in worksheets)
                {
                    if (ws.Index - 1 == outliers_array[i][0])
                    {
                        worksheet = ws;
                        break;
                    }
                }
                worksheet.Cells[row + 1, col + 1].Interior.Color = System.Drawing.Color.FromArgb(Convert.ToInt32(255 - (average_z_scores[worksheet.Index - 1][row][col] / max_weighted_z_score) * 255), 255, 255);
            }

        }


        public static void CreateCellNodesFromRange(Excel.Range range, TreeNode rangeNode, TreeNode formulaNode, TreeDict nodes)
        {
            foreach (Excel.Range cell in range)
            {
                TreeNode cellNode = null;
                //See if there is an existing node for this cell already in nodes; if there is, do not add it again - just grab the existing one
                if (!nodes.TryGetValue(ExcelParser.GetAddress(cell.Address[true, true, Excel.XlReferenceStyle.xlR1C1, false], formulaNode.getWorkbookObject(), cell.Worksheet), out cellNode))
                {
                    //TODO CORRECT THE WORKBOOK PARAMETER IN THIS LINE: (IT SHOULD BE THE WORKBOOK OF cell, WHICH SHOULD COME FROM GetReferencesFromFormula
                    var addr = ExcelParser.GetAddress(cell.Address[true, true, Excel.XlReferenceStyle.xlR1C1, false], formulaNode.getWorkbookObject(), cell.Worksheet);
                    cellNode = new TreeNode(cell.Address, cell.Worksheet, formulaNode.getWorkbookObject());
                    nodes.Add(addr, cellNode);
                }
                rangeNode.addParent(cellNode);
                cellNode.addChild(formulaNode);
                formulaNode.addParent(cellNode);
            }
        } // DoStuff ends here


        public static TreeNode MakeRangeTreeNode(TreeList ranges, Excel.Range range, TreeNode node)
        {
            TreeNode rangeNode = null;
            //See if there is an existing node for this range already in referencedRangesNodeList; if there is, do not add it again - just grab the existing one
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
                //TODO CORRECT THE WORKBOOK PARAMETER IN THIS LINE: (IT SHOULD BE THE WORKBOOK OF range, WHICH SHOULD COME FROM GetReferencesFromFormula
                rangeNode = new TreeNode(range.Address, range.Worksheet, node.getWorkbookObject());
                rangeNode.addCOM(range);
                ranges.Add(rangeNode);
            }

            return rangeNode;
        } 

    } // ConstructTree class ends here
} // namespace ends here
