using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.Office.Tools.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using System.Windows.Forms;
using System.Text.RegularExpressions;
using DataDebugMethods;

namespace DataDebug
{
    public partial class Ribbon
    {
        //private bool toolHasNotRun = true; //this is to keep track of whether the tool has already run without having cleared the colorings
        List<TreeNode> originalColorNodes = new List<TreeNode>(); //List for storing the original colors for all nodes
        List<TreeNode> nodes;        //This is a list holding all the TreeNodes in the Excel file
        TreeNode[][][] nodes_grid;   //This is a multi-dimensional array of TreeNodes that will hold all the TreeNodes -- stores the dependence graph
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

        /**
         * This is a recursive method for propagating the weights down the nodes in the tree
         * All outputs have weight 1. Their n children have weight 1/n, and so forth. 
         */
        private void propagateWeight(TreeNode node, double passed_down_weight)
        {
            if (!node.hasParents())
            {
                return;
            }
            else
            {
                int denominator = 0;  //keeps track of how many objects we are dividing the influence by
                foreach (TreeNode parent in node.getParents())
                {
                    if (parent.isRange() || parent.isChart())
                        denominator = denominator + parent.getParents().Count;
                    else
                        denominator = denominator + 1;
                }
                foreach (TreeNode parent in node.getParents())
                {
                    if (parent.isRange() || parent.isChart())
                    {
                        parent.setWeight(parent.getWeight() + passed_down_weight * parent.getParents().Count / denominator);
                        propagateWeight(parent, passed_down_weight * parent.getParents().Count / denominator);
                    }
                    else
                    {
                        parent.setWeight(parent.getWeight() + passed_down_weight / node.getParents().Count);
                        propagateWeight(parent, passed_down_weight / node.getParents().Count);
                    }
                }
            }
        }


        /**
         * This is a recursive method for propagating the weights up the nodes in the tree.
         * It is used for weighting the outputs in the computation tree -- outputs with a lot of 
         * inputs have higher weight than ones that have fewer inputs. 
         * All inputs have weight 1 and their weights get passed up in the tree and accumulated at the outputs.
         */
        private void propagateWeightUp(TreeNode node, double weight_passed_up, TreeNode originalNode)
        {
            if (!node.hasChildren())
            {
                int originalNode_row = originalNode.getWorksheetObject().Cells.get_Range(originalNode.getName()).Row - 1;
                int originalNode_col = originalNode.getWorksheetObject().Cells.get_Range(originalNode.getName()).Column - 1;
                //Mark that this output (node) is reachable from originalNode
                //Find node in output_cells
                for (int i = 0; i < output_cells.Count; i++)
                {
                    if (output_cells[i].getName().Equals(node.getName()) && output_cells[i].getWorksheet().Equals(node.getWorksheet()))
                    {
                        reachable_grid[originalNode.getWorksheetObject().Index - 1][originalNode_row][originalNode_col][i] = true;
                        reachable_impacts_grid[i].Add(new double[4] { (double)originalNode.getWorksheetObject().Index - 1, (double)originalNode_row, (double)originalNode_col, 0.0});
                        //MessageBox.Show("Output " + i + " is reachable from " + originalNode.getWorksheetObject().Name + " " + originalNode.getWorksheetObject().Cells.get_Range(originalNode.getName()).get_Address().Replace("$",""));
                        break;
                    }
                }
                return;
            }
            else
            {
                foreach (TreeNode child in node.getChildren())
                {
                    child.setWeight(child.getWeight() + weight_passed_up);
                    propagateWeightUp(child, 1.0, originalNode);
                }
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
            TimeSpan impact_scoring_timespan = global_stopwatch.Elapsed;
            TimeSpan swapping_timespan = global_stopwatch.Elapsed;
            input_cells_in_computation_count = 0;
            raw_input_cells_in_computation_count = 0;

            // Get a range representing the formula cells for each worksheet in each workbook
            var analysisRanges = ConstructTree.GetFormulaRanges(Globals.ThisAddIn.Application.Worksheets, Globals.ThisAddIn.Application);
            formula_cells_count = ConstructTree.CountFormulaCells(analysisRanges);
            
            //First we create nodes for every non-null cell; then we will operate on these node objects, connecting them in the tree, etc. 
            //This includes cells that contain constants and formulas
            //Go through every worksheet
            foreach (Excel.Range worksheet_range in analysisRanges)
            {
                // Go through every cell of every worksheet
                if (worksheet_range != null)
                {
                    foreach (Excel.Range cell in worksheet_range)
                    {
                        if (cell.Value != null)
                        {
                            TreeNode n = new TreeNode(cell.Address, cell.Worksheet.Name, Globals.ThisAddIn.Application.ActiveWorkbook);  //Create a TreeNode for every cell with the name being the cell's address and set the node's worksheet appropriately
                            if (cell.HasFormula)
                            {
                                n.setIsFormula();
                            }
                            try
                            {
                                nodes_grid[cell.Worksheet.Index - 1][cell.Row - 1][cell.Column - 1] = n;
                            }
                            catch
                            {
                                cell.Interior.Color = System.Drawing.Color.Purple;
                            }
                        }
                    }
                }
                else
                {
                    continue;
                }
            }

            //Next we go through the cells that contain formulas in order to extract the dependencies between them and their inputs
            //For every cell that contains a formula, we get the node we created for that cell. Then we parse the formula using a regular expresion 
            //to find any references to cells or ranges. (We first look for references to ranges, because they supersede the single cell references.)
            //Whenever a reference is found, we update the parent-child relationship between the formula cell and the referenced cell or range.
            //If a range reference is found, we create a node representing that range, and we also create nodes for all of the cells that compose it. 
            //The range is connected to the formula cell, and the composing cells are connected to the range. 
            //If a single cell reference is found, we connect it to the formula cell directly. 
            
            // Get the names of all worksheets in the workbook and store them in the array worksheet_names
            String[] worksheet_names = new String[Globals.ThisAddIn.Application.Worksheets.Count]; // Array holding the names of every worksheet in the workbook
            int index_worksheet_names = 0; // Index for populating the worksheet_names
            foreach (Excel.Worksheet worksheet in Globals.ThisAddIn.Application.Worksheets)
            {
                worksheet_names[index_worksheet_names] = worksheet.Name;
                index_worksheet_names++;
            }

            foreach (Excel.Range worksheet_range in analysisRanges)
            {
                if (worksheet_range != null) //if the worksheet is not blank, analyze its contents
                {
                    foreach (Excel.Range c in worksheet_range)
                    {
                        if (c.HasFormula)
                        {
                            TreeNode formula_cell = null;
                            //Look for the node object for the current cell in the existing TreeNodes
                            //Check if this cell's coordinates are within the bounds of the used range, otherwise there will be an index out of bounds error
                            if (c.Column <= (c.Worksheet.UsedRange.Columns.Count + c.Worksheet.UsedRange.Column) && c.Row <= (c.Worksheet.UsedRange.Rows.Count + c.Worksheet.UsedRange.Row))
                            {
                                //if a TreeNode exists for this cell already
                                if (nodes_grid[c.Worksheet.Index - 1][c.Row - 1][c.Column - 1] != null)
                                {
                                    formula_cell = nodes_grid[c.Worksheet.Index - 1][c.Row - 1][c.Column - 1];
                                }
                            }
                            else
                            {
                                //Create a tree node for the cell
                                TreeNode n = new TreeNode(c.Address, c.Worksheet.Name, Globals.ThisAddIn.Application.ActiveWorkbook);  //Create a TreeNode for every cell with the name being the cell's address and set the node's worksheet appropriately
                                if (c.HasFormula)
                                {
                                    n.setIsFormula();
                                }
                                try
                                {
                                    nodes_grid[c.Worksheet.Index - 1][c.Row - 1][c.Column - 1] = n;
                                }
                                catch
                                {
                                    c.Interior.Color = System.Drawing.Color.Purple;
                                }
                                formula_cell = n;
                            }

                            string formula = c.Formula;  //The formula contained in the cell
                            //if (formula.Contains("HLOOKUP") || formula.Contains("VLOOKUP"))
                            //{
                            //    continue;
                            //}
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
                            MatchCollection matchedRanges = null;
                            MatchCollection matchedCells = null;
                            int ws_index = 1;
                            foreach (string s in worksheet_names)
                            {
                                string worksheet_name = s.Replace("+", @"\+").Replace("^", @"\^").Replace("$", @"\$").Replace(".", @"\."); //Escape certain characters in the regular expression
                                //First look for range references of the form 'worksheet_name'!A1:A10 in the formula (with quotation marks around the name)
                                if (toggle_compile_regex.Checked)
                                {
                                    //Regex regex = new Regex(@"('" + worksheet_name + @"'!\$?[A-Z]+\$?[1-9]\d*:\$?[A-Z]+\$?[1-9]\d*)", RegexOptions.Compiled);
                                    //Regex regex = regex_array[ws_index - 1];
                                    matchedRanges = regex_array[4*(ws_index-1)].Matches(formula);
                                }
                                else
                                {
                                    matchedRanges = Regex.Matches(formula, @"('" + worksheet_name + @"'!\$?[A-Z]+\$?[1-9]\d*:\$?[A-Z]+\$?[1-9]\d*)");  //A collection of all the references in the formula to ranges in the particular worksheet; each item is a range reference of the form 'worksheet_name'!A1:A10
                                }
                                foreach (Match match in matchedRanges)
                                {
                                    formula = formula.Replace(match.Value, "");
                                    string ws_name = worksheet_name; // match.Value.Substring(1, match.Value.LastIndexOf("!") - 2); // Get the name of the worksheet being referenced
                                    string range_coordinates = match.Value.Substring(match.Value.LastIndexOf("!") + 1); //match.Value.Replace("'" + ws_name + "'!", "");
                                    string[] endCells = range_coordinates.Split(':');     //Split up each matched range into the start and end cells of the range
                                    TreeNode range = null;
                                    //Try to find the range in existing TreeNodes
                                    foreach (TreeNode n in ranges)
                                    {
                                        if (n.getName().Replace("$", "") == endCells[0].Replace("$", "") + "_to_" + endCells[1].Replace("$", "") && n.getWorksheet() == ws_name)
                                        {
                                            //MessageBox.Show("Found node:" + ws_name + "_" + endCells[0] + ":" + endCells[1]);
                                            range = n;
                                        }
                                        else
                                        {
                                            continue;
                                        }
                                    }
                                        
                                    //If it was not found, create it
                                    if (range == null)
                                    {
                                        range = new TreeNode(endCells[0].Replace("$", "") + "_to_" + endCells[1].Replace("$", ""), ws_name, Globals.ThisAddIn.Application.ActiveWorkbook);
                                        //MessageBox.Show("Created range node:" + ws_name + "_" + endCells[0] + ":" + endCells[1]);
                                        ranges.Add(range);
                                    }
                                    formula_cell.addParent(range);
                                    range.addChild(formula_cell);
                                    //Add each cell contained in the range to the dependencies
                                    foreach (Excel.Range cellInRange in Globals.ThisAddIn.Application.Worksheets[ws_index].Range[endCells[0], endCells[1]])
                                    {
                                        TreeNode input_cell = null;
                                        //Find the node object for the current cell in the existing TreeNodes
                                        //Check if this cell's coordinates are within the bounds of the used range, otherwise there will be an index out of bounds error
                                        if (cellInRange.Column <= (cellInRange.Worksheet.UsedRange.Columns.Count + cellInRange.Worksheet.UsedRange.Column) && cellInRange.Row <= (cellInRange.Worksheet.UsedRange.Rows.Count + cellInRange.Worksheet.UsedRange.Row))
                                        {
                                            //if a TreeNode exists for this cell already
                                            if (nodes_grid[cellInRange.Worksheet.Index - 1][cellInRange.Row - 1][cellInRange.Column - 1] != null)
                                            {
                                                input_cell = nodes_grid[cellInRange.Worksheet.Index - 1][cellInRange.Row - 1][cellInRange.Column - 1];
                                            }
                                        }
                                            
                                        //If it wasn't found, then it is blank, and we have to create a TreeNode for it
                                        if (input_cell == null)
                                        {
                                            input_cell = new TreeNode(cellInRange.Address, cellInRange.Worksheet.Name, Globals.ThisAddIn.Application.ActiveWorkbook);
                                            //Check if this cell's coordinates are within the bounds of the used range, otherwise there will be an index out of bounds error
                                            if (cellInRange.Column <= (cellInRange.Worksheet.UsedRange.Columns.Count + cellInRange.Worksheet.UsedRange.Column) && cellInRange.Row <= (cellInRange.Worksheet.UsedRange.Rows.Count + cellInRange.Worksheet.UsedRange.Row))
                                            {
                                                nodes_grid[cellInRange.Worksheet.Index - 1][cellInRange.Row - 1][cellInRange.Column - 1] = input_cell;
                                            }
                                        }

                                        //Update the dependencies
                                        range.addParent(input_cell);
                                        input_cell.addChild(range);
                                    }
                                }

                                //Next look for range references of the form worksheet_name!A1:A10 in the formula (no quotation marks around the name)
                                if (toggle_compile_regex.Checked)
                                {
                                    //Regex regex = new Regex(@"(" + worksheet_name + @"!\$?[A-Z]+\$?[1-9]\d*:\$?[A-Z]+\$?[1-9]\d*)", RegexOptions.Compiled);
                                    //Regex regex = regex_array[ws_index + 1 - 1];
                                    matchedRanges = regex_array[4*(ws_index - 1) + 1].Matches(formula);
                                }
                                else
                                {
                                    matchedRanges = Regex.Matches(formula, @"(" + worksheet_name + @"!\$?[A-Z]+\$?[1-9]\d*:\$?[A-Z]+\$?[1-9]\d*)");  //A collection of all the range references in the formula; each item is a range reference such as A1:A10
                                }
                                
                                foreach (Match match in matchedRanges)
                                {
                                    formula = formula.Replace(match.Value, "");
                                    string ws_name = worksheet_name; // match.Value.Substring(0, match.Value.LastIndexOf("!")); // Get the name of the worksheet being referenced
                                    string range_coordinates = match.Value.Substring(match.Value.LastIndexOf("!") + 1);  //match.Value.Replace(ws_name + "!", "");
                                    string[] endCells = range_coordinates.Split(':');     //Split up each matched range into the start and end cells of the range
                                    TreeNode range = null;
                                    //Try to find the range in existing TreeNodes
                                    foreach (TreeNode n in ranges)
                                    {
                                        if (n.getName().Replace("$", "") == endCells[0].Replace("$", "") + "_to_" + endCells[1].Replace("$", "") && n.getWorksheet() == ws_name)
                                        {
                                            //MessageBox.Show("Found node:" + ws_name + "_" + endCells[0] + ":" + endCells[1]);
                                            range = n;
                                        }
                                        else
                                        {
                                            continue;
                                        }
                                    }
                                    //If it does not exist, create it
                                    if (range == null)
                                    {
                                        range = new TreeNode(endCells[0].Replace("$", "") + "_to_" + endCells[1].Replace("$", ""), ws_name, Globals.ThisAddIn.Application.ActiveWorkbook);
                                        //MessageBox.Show("Created node:" + ws_name + "_" + endCells[0] + ":" + endCells[1]);
                                        ranges.Add(range);
                                    }
                                    formula_cell.addParent(range);
                                    range.addChild(formula_cell);
                                    //Add each cell contained in the range to the dependencies
                                    foreach (Excel.Range cellInRange in Globals.ThisAddIn.Application.Worksheets[ws_index].Range[endCells[0], endCells[1]])
                                    {
                                        TreeNode input_cell = null;
                                        //Find the node object for the current cell in the existing TreeNodes
                                        //Check if this cell's coordinates are within the bounds of the used range, otherwise there will be an index out of bounds error
                                        if (cellInRange.Column <= (cellInRange.Worksheet.UsedRange.Columns.Count + cellInRange.Worksheet.UsedRange.Column) && cellInRange.Row <= (cellInRange.Worksheet.UsedRange.Rows.Count + cellInRange.Worksheet.UsedRange.Row))
                                        {
                                            //if a TreeNode exists for this cell already
                                            if (nodes_grid[cellInRange.Worksheet.Index - 1][cellInRange.Row - 1][cellInRange.Column - 1] != null)
                                            {
                                                input_cell = nodes_grid[cellInRange.Worksheet.Index - 1][cellInRange.Row - 1][cellInRange.Column - 1];
                                            }
                                        }
                                        //If it wasn't found, then it is blank, and we have to create a TreeNode for it
                                        if (input_cell == null)
                                        {
                                            input_cell = new TreeNode(cellInRange.Address, cellInRange.Worksheet.Name, Globals.ThisAddIn.Application.ActiveWorkbook);
                                            //Check if this cell's coordinates are within the bounds of the used range, otherwise there will be an index out of bounds error
                                            if (cellInRange.Column <= (cellInRange.Worksheet.UsedRange.Columns.Count + cellInRange.Worksheet.UsedRange.Column) && cellInRange.Row <= (cellInRange.Worksheet.UsedRange.Rows.Count + cellInRange.Worksheet.UsedRange.Row))
                                            {
                                                nodes_grid[cellInRange.Worksheet.Index - 1][cellInRange.Row - 1][cellInRange.Column - 1] = input_cell;
                                            }
                                        }

                                        //Update the dependencies
                                        range.addParent(input_cell);
                                        input_cell.addChild(range);
                                    }
                                }

                                // Now we look for references of the kind 'worksheet_name'!A1 (with quotation marks)
                                if (toggle_compile_regex.Checked)
                                {
                                    //Regex regex = new Regex(@"('" + worksheet_name + @"'!\$?[A-Z]+\$?[1-9]\d*)", RegexOptions.Compiled);
                                    //Regex regex = regex_array[ws_index + 2 - 1];
                                    matchedCells = regex_array[4*(ws_index - 1) + 2].Matches(formula);
                                }
                                else
                                {
                                    matchedCells = Regex.Matches(formula, @"('" + worksheet_name + @"'!\$?[A-Z]+\$?[1-9]\d*)"); //matchedCells is a collection of all the references in the formula to cells in the specific worksheet, where the reference has the form 'worksheet_name'!A1
                                }
                                foreach (Match match in matchedCells)
                                {
                                    formula = formula.Replace(match.Value, "");
                                    string ws_name = worksheet_name; // match.Value.Substring(1, match.Value.LastIndexOf("!") - 2); // Get the name of the worksheet being referenced
                                    string cell_coordinates = match.Value.Substring(match.Value.LastIndexOf("!") + 1);
                                    //Get the actual cell that is being referenced
                                    Excel.Range input = null;
                                    foreach (Excel.Worksheet ws in Globals.ThisAddIn.Application.Worksheets)
                                    {
                                        //Find the worksheet object that the match belongs to
                                        if (ws.Name == ws_name)
                                        {
                                            input = ws.get_Range(cell_coordinates);
                                        }
                                    }
                                    TreeNode input_cell = null;
                                    //Find the node object for the current cell in the existing TreeNodes
                                    //Check if this cell's coordinates are within the bounds of the used range of its spreadsheet, otherwise there will be an index out of bounds error
                                    if (input.Column <= (input.Worksheet.UsedRange.Columns.Count + input.Worksheet.UsedRange.Column) && input.Row <= (input.Worksheet.UsedRange.Rows.Count + input.Worksheet.UsedRange.Row))
                                    {
                                        //if a TreeNode exists for this cell already, use it
                                        if (nodes_grid[input.Worksheet.Index - 1][input.Row - 1][input.Column - 1] != null)
                                        {
                                            input_cell = nodes_grid[input.Worksheet.Index - 1][input.Row - 1][input.Column - 1];
                                        }
                                    }
                                    //If it wasn't found, then it is blank, and we have to create a TreeNode for it
                                    if (input_cell == null)
                                    {
                                        input_cell = new TreeNode(cell_coordinates.Replace("$", ""), ws_name, Globals.ThisAddIn.Application.ActiveWorkbook);
                                        //Check if this cell's coordinates are within the bounds of the used range, otherwise there will be an index out of bounds error
                                        if (input.Column <= (input.Worksheet.UsedRange.Columns.Count + input.Worksheet.UsedRange.Column) && input.Row <= (input.Worksheet.UsedRange.Rows.Count + input.Worksheet.UsedRange.Row))
                                        {
                                            nodes_grid[input.Worksheet.Index - 1][input.Row - 1][input.Column - 1] = input_cell;
                                        }
                                    }

                                    //Update the dependencies
                                    formula_cell.addParent(input_cell);
                                    input_cell.addChild(formula_cell);
                                }

                                //Lastly we look for references of the kind worksheet_name!A1 (without quotation marks)
                                if (toggle_compile_regex.Checked)
                                {
                                    //Regex regex = new Regex(@"(" + worksheet_name + @"!\$?[A-Z]+\$?[1-9]\d*)", RegexOptions.Compiled);
                                    //Regex regex = regex_array[ws_index + 3 - 1];
                                    matchedCells = regex_array[4*(ws_index - 1) + 3].Matches(formula);
                                }
                                else
                                {
                                    matchedCells = Regex.Matches(formula, @"(" + worksheet_name + @"!\$?[A-Z]+\$?[1-9]\d*)");
                                }
                                foreach (Match match in matchedCells)
                                {
                                    formula = formula.Replace(match.Value, "");
                                    string ws_name = worksheet_name; //match.Value.Substring(0, match.Value.LastIndexOf("!")); // Get the name of the worksheet being referenced
                                    string cell_coordinates = match.Value.Substring(match.Value.LastIndexOf("!") + 1);
                                    //MessageBox.Show(formula_cell.getName() + " refers to the cell " + ws_name + "!" + cell_coordinates);
                                    //Get the actual cell that is being referenced
                                    Excel.Range input = null;
                                    foreach (Excel.Worksheet ws in Globals.ThisAddIn.Application.Worksheets)
                                    {
                                        //Find the worksheet object that the match belongs to
                                        if (ws.Name == ws_name)
                                        {
                                            input = ws.get_Range(cell_coordinates);
                                        }
                                    }
                                    TreeNode input_cell = null;
                                    //Find the node object for the current cell in the existing TreeNodes
                                    //Check if this cell's coordinates are within the bounds of the used range of its spreadsheet, otherwise there will be an index out of bounds error
                                    if (input.Column <= (input.Worksheet.UsedRange.Columns.Count + input.Worksheet.UsedRange.Column) && input.Row <= (input.Worksheet.UsedRange.Rows.Count + input.Worksheet.UsedRange.Row))
                                    {
                                        //if a TreeNode exists for this cell already, use it
                                        if (nodes_grid[input.Worksheet.Index - 1][input.Row - 1][input.Column - 1] != null)
                                        {
                                            input_cell = nodes_grid[input.Worksheet.Index - 1][input.Row - 1][input.Column - 1];
                                        }
                                    }
                                    //If it wasn't found, then it is blank, and we have to create a TreeNode for it
                                    if (input_cell == null)
                                    {
                                        input_cell = new TreeNode(cell_coordinates.Replace("$", ""), ws_name, Globals.ThisAddIn.Application.ActiveWorkbook);
                                        //Check if this cell's coordinates are within the bounds of the used range, otherwise there will be an index out of bounds error
                                        if (input.Column <= (input.Worksheet.UsedRange.Columns.Count + input.Worksheet.UsedRange.Column) && input.Row <= (input.Worksheet.UsedRange.Rows.Count + input.Worksheet.UsedRange.Row))
                                        {
                                            nodes_grid[input.Worksheet.Index - 1][input.Row - 1][input.Column - 1] = input_cell;
                                        }
                                    }

                                    //Update the dependencies
                                    formula_cell.addParent(input_cell);
                                    input_cell.addChild(formula_cell);
                                }
                                ws_index++;
                            }

                            string patternRange = @"(\$?[A-Z]+\$?[1-9]\d*:\$?[A-Z]+\$?[1-9]\d*)";  //Regex for matching range references in formulas such as A1:A10, or $A$1:$A$10 etc.
                            string patternCell = @"(\$?[A-Z]+\$?[1-9]\d*)";        //Regex for matching single cell references such as A1 or $A$1, etc. 

                            //First look for range references in the formula
                            if (toggle_compile_regex.Checked)
                            {
                                //Regex regex = new Regex(patternRange, RegexOptions.Compiled);
                                //Regex regex = regex_array[regex_array.Length - 2];
                                matchedRanges = regex_array[regex_array.Length - 2].Matches(formula);
                            }
                            else
                            {
                                matchedRanges = Regex.Matches(formula, patternRange);  //A collection of all the range references in the formula; each item is a range reference such as A1:A10
                            }
                            //List<Excel.Range> rangeList = new List<Excel.Range>();
                            foreach (Match match in matchedRanges)
                            {
                                formula = formula.Replace(match.Value, "");
                                string[] endCells = match.Value.Split(':');     //Split up each matched range into the start and end cells of the range
                                TreeNode range = null;
                                //Try to find the range in existing TreeNodes
                                foreach (TreeNode n in ranges)
                                {
                                    if (n.getName().Replace("$", "") == endCells[0].Replace("$", "") + "_to_" + endCells[1].Replace("$", "") && n.getWorksheet() == c.Worksheet.Name)
                                    {
                                        range = n;
                                    }
                                    else
                                    {
                                        continue;
                                    }
                                }
                                //If it does not exist, create it
                                if (range == null)
                                {
                                    //MessageBox.Show("Created range node:" + c.Worksheet.Name + "_" + endCells[0] + ":" + endCells[1]);
                                    range = new TreeNode(endCells[0].Replace("$", "") + "_to_" + endCells[1].Replace("$", ""), c.Worksheet.Name, Globals.ThisAddIn.Application.ActiveWorkbook);
                                    ranges.Add(range);
                                }
                                formula_cell.addParent(range);
                                range.addChild(formula_cell);
                                //Add each cell contained in the range to the dependencies
                                foreach (Excel.Range cellInRange in c.Worksheet.Range[endCells[0], endCells[1]])
                                {
                                    TreeNode input_cell = null;
                                    //Find the node object for the current cell in the existing TreeNodes
                                    //Check if this cell's coordinates are within the bounds of the used range, otherwise there will be an index out of bounds error
                                    if (cellInRange.Column <= (cellInRange.Worksheet.UsedRange.Columns.Count + cellInRange.Worksheet.UsedRange.Row) && cellInRange.Row <= (cellInRange.Worksheet.UsedRange.Rows.Count + cellInRange.Worksheet.UsedRange.Row))
                                    {
                                        //if a TreeNode exists for this cell already, use it
                                        if (nodes_grid[cellInRange.Worksheet.Index - 1][cellInRange.Row - 1][cellInRange.Column - 1] != null)
                                        {
                                            input_cell = nodes_grid[cellInRange.Worksheet.Index - 1][cellInRange.Row - 1][cellInRange.Column - 1];
                                        }
                                    }
                                    //If it wasn't found, then it is blank, and we have to create a TreeNode for it
                                    if (input_cell == null)
                                    {
                                        input_cell = new TreeNode(cellInRange.Address, cellInRange.Worksheet.Name, Globals.ThisAddIn.Application.ActiveWorkbook);
                                        //Check if this cell's coordinates are within the bounds of the used range, otherwise there will be an index out of bounds error
                                        if (cellInRange.Column <= (cellInRange.Worksheet.UsedRange.Columns.Count + cellInRange.Worksheet.UsedRange.Column) && cellInRange.Row <= (cellInRange.Worksheet.UsedRange.Rows.Count + cellInRange.Worksheet.UsedRange.Row))
                                        {
                                            nodes_grid[cellInRange.Worksheet.Index - 1][cellInRange.Row - 1][cellInRange.Column - 1] = input_cell;
                                        }
                                    }

                                    //Update the dependencies
                                    range.addParent(input_cell);
                                    input_cell.addChild(range);
                                }
                            }

                            //Find any references to named ranges
                            //TODO -- this should probably be done in a better way - with a regular expression that will catch things like this:
                            //"+range_name", "-range_name", "*range_name", etc., because right now a range name may be part of the name of a 
                            //formula that is used. For instance a range could be named "s", and if the formula has the "sum" function in it, we will 
                            //falsely detect a reference to "s". This does not affect the correctness of the algorithm, because all we care about 
                            //from the dependence graph is identifying which cells are outputs, and identifying user-defined ranges
                            //and this type of error will not affect either one
                            foreach (Excel.Name named_range in Globals.ThisAddIn.Application.Names)
                            {
                                if (formula.Contains(named_range.Name))
                                {
                                    formula = formula.Replace(named_range.Name, "");
                                }
                                else
                                {
                                    continue;
                                }
                                //If this named range holds a range
                                if (named_range.RefersToRange.Address.Contains(":"))
                                {
                                    string[] endCells = named_range.RefersToRange.Address.Split(':');     //Split up each named range into the start and end cells of the range
                                    TreeNode range = null;
                                    //Try to find the range in existing TreeNodes
                                    foreach (TreeNode n in ranges)
                                    {
                                        if (n.getName().Replace("$", "") == endCells[0].Replace("$", "") + "_to_" + endCells[1].Replace("$", "") && n.getWorksheet() == named_range.RefersToRange.Worksheet.Name)
                                        {
                                            range = n;
                                        }
                                        else
                                        {
                                            continue;
                                        }
                                    }
                                    //If it does not exist, create it
                                    if (range == null)
                                    {
                                        //MessageBox.Show("Created range node:" + c.Worksheet.Name + "_" + endCells[0] + ":" + endCells[1]);
                                        range = new TreeNode(endCells[0].Replace("$", "") + "_to_" + endCells[1].Replace("$", ""), named_range.RefersToRange.Worksheet.Name, Globals.ThisAddIn.Application.ActiveWorkbook);
                                        ranges.Add(range);
                                    }
                                    formula_cell.addParent(range);
                                    range.addChild(formula_cell);
                                    //Add each cell contained in the range to the dependencies
                                    foreach (Excel.Range cellInRange in named_range.RefersToRange.Worksheet.Range[endCells[0], endCells[1]])
                                    {
                                        TreeNode input_cell = null;
                                        //Find the node object for the current cell in the existing TreeNodes
                                        //Check if this cell's coordinates are within the bounds of the used range, otherwise there will be an index out of bounds error
                                        if (cellInRange.Column <= (cellInRange.Worksheet.UsedRange.Columns.Count + cellInRange.Worksheet.UsedRange.Column) && cellInRange.Row <= (cellInRange.Worksheet.UsedRange.Rows.Count + cellInRange.Worksheet.UsedRange.Row))
                                        {
                                            //if a TreeNode exists for this cell already, use it
                                            if (nodes_grid[cellInRange.Worksheet.Index - 1][cellInRange.Row - 1][cellInRange.Column - 1] != null)
                                            {
                                                input_cell = nodes_grid[cellInRange.Worksheet.Index - 1][cellInRange.Row - 1][cellInRange.Column - 1];
                                            }
                                        }
                                        //If it wasn't found, then it is blank, and we have to create a TreeNode for it
                                        if (input_cell == null)
                                        {
                                            input_cell = new TreeNode(cellInRange.Address, cellInRange.Worksheet.Name, Globals.ThisAddIn.Application.ActiveWorkbook);
                                            //Check if this cell's coordinates are within the bounds of the used range, otherwise there will be an index out of bounds error
                                            if (cellInRange.Column <= (cellInRange.Worksheet.UsedRange.Columns.Count + cellInRange.Worksheet.UsedRange.Column) && cellInRange.Row <= (cellInRange.Worksheet.UsedRange.Rows.Count + cellInRange.Worksheet.UsedRange.Row))
                                            {
                                                nodes_grid[cellInRange.Worksheet.Index - 1][cellInRange.Row - 1][cellInRange.Column - 1] = input_cell;
                                            }
                                        }

                                        //Update the dependencies
                                        range.addParent(input_cell);
                                        input_cell.addChild(range);
                                    }
                                }
                                else  //If this named range holds a cell
                                {
                                    Excel.Range input =  named_range.RefersToRange;
                                    TreeNode input_cell = null;
                                    //Find the node object for the current cell in the existing TreeNodes
                                    //Check if this cell's coordinates are within the bounds of the used range, otherwise there will be an index out of bounds error
                                    if (input.Column <= (input.Worksheet.UsedRange.Columns.Count + input.Worksheet.UsedRange.Column) && input.Row <= (input.Worksheet.UsedRange.Rows.Count + input.Worksheet.UsedRange.Row))
                                    {
                                        //if a TreeNode exists for this cell already, use it
                                        if (nodes_grid[input.Worksheet.Index - 1][input.Row - 1][input.Column - 1] != null)
                                        {
                                            input_cell = nodes_grid[input.Worksheet.Index - 1][input.Row - 1][input.Column - 1];
                                        }
                                    }
                                    //If it wasn't found, then it is blank, and we have to create a TreeNode for it
                                    if (input_cell == null)
                                    {
                                        input_cell = new TreeNode(named_range.RefersToRange.Address.Replace("$", ""), named_range.RefersToRange.Worksheet.Name, Globals.ThisAddIn.Application.ActiveWorkbook);
                                        //Check if this cell's coordinates are within the bounds of the used range, otherwise there will be an index out of bounds error
                                        if (input.Column <= (input.Worksheet.UsedRange.Columns.Count + input.Worksheet.UsedRange.Column) && input.Row <= (input.Worksheet.UsedRange.Rows.Count + input.Worksheet.UsedRange.Row))
                                        {
                                            nodes_grid[input.Worksheet.Index - 1][input.Row - 1][input.Column - 1] = input_cell;
                                        }
                                    }
                                    //Update the dependencies
                                    formula_cell.addParent(input_cell);
                                    input_cell.addChild(formula_cell);
                                }
                            }

                            if (toggle_compile_regex.Checked)
                            {
                                //Regex regex = new Regex(patternCell, RegexOptions.Compiled);
                                //Regex regex = regex_array[regex_array.Length - 1];
                                matchedCells = regex_array[regex_array.Length - 1].Matches(formula);
                            }
                            else
                            {
                                matchedCells = Regex.Matches(formula, patternCell);  //matchedCells is a collection of all the cells that are referenced by the formula
                            }
                            foreach (Match m in matchedCells)
                            {
                                Excel.Range input = c.Worksheet.get_Range(m.Value);
                                TreeNode input_cell = null;
                                //MessageBox.Show(m.Value);
                                //Find the node object for the current cell in the existing TreeNodes
                                //Check if this cell's coordinates are within the bounds of the used range, otherwise there will be an index out of bounds error
                                if (input.Column <= (input.Worksheet.UsedRange.Columns.Count + input.Worksheet.UsedRange.Column) && input.Row <= (input.Worksheet.UsedRange.Rows.Count + input.Worksheet.UsedRange.Row))
                                {
                                    //if a TreeNode exists for this cell already, use it
                                    if (nodes_grid[input.Worksheet.Index - 1][input.Row - 1][input.Column - 1] != null)
                                    {
                                        input_cell = nodes_grid[input.Worksheet.Index - 1][input.Row - 1][input.Column - 1];
                                    }
                                }
                                //If it wasn't found, then it is blank, and we have to create a TreeNode for it
                                if (input_cell == null)
                                {
                                    input_cell = new TreeNode(m.Value.Replace("$", ""), c.Worksheet.Name, Globals.ThisAddIn.Application.ActiveWorkbook);
                                    //Check if this cell's coordinates are within the bounds of the used range, otherwise there will be an index out of bounds error
                                    if (input.Column <= (input.Worksheet.UsedRange.Columns.Count + input.Worksheet.UsedRange.Column) && input.Row <= (input.Worksheet.UsedRange.Rows.Count + input.Worksheet.UsedRange.Row))
                                    {
                                        nodes_grid[input.Worksheet.Index - 1][input.Row - 1][input.Column - 1] = input_cell;
                                    }
                                }
                                //Update the dependencies
                                formula_cell.addParent(input_cell);
                                input_cell.addChild(formula_cell);
                            } 
                        }
                    }
                }
                else  // If this worksheet is blank, move on to the next one
                {
                    continue;
                }
            }

            //Print out text for GraphViz representation of the dependence graph
            //string tree1 = "";
            //foreach (TreeNode node in nodes)
            //{
            //    tree1 += node.toGVString(0) + "\n";
            //}
            //Display disp1 = new Display();
            //disp1.textBox1.Text = "digraph g{" + tree1 + "}";
            //disp1.ShowDialog();
            foreach (Excel.Chart chart in Globals.ThisAddIn.Application.Charts)
            {
                //TODO The naming convention for TreeNode charts is kind of a hack; could fail if two charts have the same names when white spaces are removed - maybe add a random hash at the end
                TreeNode chart_node = new TreeNode(chart.Name, "none", Globals.ThisAddIn.Application.ActiveWorkbook);
                chart_node.setChart(true);
                charts.Add(chart_node);
                foreach (Excel.Series series in (Excel.SeriesCollection)chart.SeriesCollection(Type.Missing))
                {
                    string formula = series.Formula;  //The formula contained in the cell
                    //find_references(chart_node, formula);
                    
                    MatchCollection matchedRanges = null;
                    MatchCollection matchedCells = null;
                    int ws_index = 1;
                    foreach (string s in worksheet_names)
                    {
                        string worksheet_name = s.Replace("+", @"\+").Replace("^", @"\^").Replace("$", @"\$").Replace(".", @"\."); //Escape certain characters in the regular expression
                        //First look for range references of the form 'worksheet_name'!A1:A10 in the formula (with quotation marks around the name)
                        if (toggle_compile_regex.Checked)
                        {
                            //Regex regex = new Regex(@"('" + worksheet_name + @"'!\$?[A-Z]+\$?[1-9]\d*:\$?[A-Z]+\$?[1-9]\d*)", RegexOptions.Compiled);
                            //Regex regex = regex_array[ws_index - 1];
                            matchedRanges = regex_array[4*(ws_index - 1)].Matches(formula);
                        }
                        else
                        {
                            matchedRanges = Regex.Matches(formula, @"('" + worksheet_name + @"'!\$?[A-Z]+\$?[1-9]\d*:\$?[A-Z]+\$?[1-9]\d*)");  //A collection of all the references in the formula to ranges in the particular worksheet; each item is a range reference of the form 'worksheet_name'!A1:A10
                        }
                        foreach (Match match in matchedRanges)
                        {
                            formula = formula.Replace(match.Value, "");
                            string ws_name = worksheet_name; // match.Value.Substring(1, match.Value.LastIndexOf("!") - 2); // Get the name of the worksheet being referenced
                            string range_coordinates = match.Value.Substring(match.Value.LastIndexOf("!") + 1); //match.Value.Replace("'" + ws_name + "'!", "");
                            string[] endCells = range_coordinates.Split(':');     //Split up each matched range into the start and end cells of the range
                            TreeNode range = null;
                            //Try to find the range in existing TreeNodes
                            foreach (TreeNode n in ranges)
                            {
                                if (n.getName().Replace("$", "") == endCells[0].Replace("$", "") + "_to_" + endCells[1].Replace("$", "") && n.getWorksheet() == ws_name)
                                {
                                    //MessageBox.Show("Found node:" + ws_name + "_" + endCells[0] + ":" + endCells[1]);
                                    range = n;
                                }
                                else
                                {
                                    continue;
                                }
                            }
                            //If it was not found, create it
                            if (range == null)
                            {
                                range = new TreeNode(endCells[0].Replace("$", "") + "_to_" + endCells[1].Replace("$", ""), ws_name, Globals.ThisAddIn.Application.ActiveWorkbook);
                                //MessageBox.Show("Created range node:" + ws_name + "_" + endCells[0] + ":" + endCells[1]);
                                ranges.Add(range);
                            }
                            chart_node.addParent(range);
                            range.addChild(chart_node);
                            //Add each cell contained in the range to the dependencies
                            foreach (Excel.Range cellInRange in Globals.ThisAddIn.Application.Worksheets[ws_index].Range[endCells[0], endCells[1]])
                            {
                                TreeNode input_cell = null;
                                //Find the node object for the current cell in the existing TreeNodes
                                //Check if this cell's coordinates are within the bounds of the used range, otherwise there will be an index out of bounds error
                                if (cellInRange.Column <= (cellInRange.Worksheet.UsedRange.Columns.Count + cellInRange.Worksheet.UsedRange.Column) && cellInRange.Row <= (cellInRange.Worksheet.UsedRange.Rows.Count + cellInRange.Worksheet.UsedRange.Row))
                                {
                                    //if a TreeNode exists for this cell already
                                    if (nodes_grid[cellInRange.Worksheet.Index - 1][cellInRange.Row - 1][cellInRange.Column - 1] != null)
                                    {
                                        input_cell = nodes_grid[cellInRange.Worksheet.Index - 1][cellInRange.Row - 1][cellInRange.Column - 1];
                                    }
                                }
                                //If it wasn't found, then it is blank, and we have to create a TreeNode for it
                                if (input_cell == null)
                                {
                                    input_cell = new TreeNode(cellInRange.Address, cellInRange.Worksheet.Name, Globals.ThisAddIn.Application.ActiveWorkbook);
                                    //Check if this cell's coordinates are within the bounds of the used range, otherwise there will be an index out of bounds error
                                    if (cellInRange.Column <= (cellInRange.Worksheet.UsedRange.Columns.Count + cellInRange.Worksheet.UsedRange.Column) && cellInRange.Row <= (cellInRange.Worksheet.UsedRange.Rows.Count + cellInRange.Worksheet.UsedRange.Row))
                                    {
                                        nodes_grid[cellInRange.Worksheet.Index - 1][cellInRange.Row - 1][cellInRange.Column - 1] = input_cell;
                                    }
                                }

                                //Update the dependencies
                                range.addParent(input_cell);
                                input_cell.addChild(range);
                            }
                        }

                        //Next look for range references of the form worksheet_name!A1:A10 in the formula (no quotation marks around the name)
                        if (toggle_compile_regex.Checked)
                        {
                            //Regex regex = new Regex(@"('" + worksheet_name + @"'!\$?[A-Z]+\$?[1-9]\d*:\$?[A-Z]+\$?[1-9]\d*)", RegexOptions.Compiled);
                            //Regex regex = regex_array[ws_index - 1];
                            matchedRanges = regex_array[4*(ws_index - 1) + 1].Matches(formula);
                        }
                        else
                        {
                            matchedRanges = Regex.Matches(formula, @"(" + worksheet_name + @"!\$?[A-Z]+\$?[1-9]\d*:\$?[A-Z]+\$?[1-9]\d*)");  //A collection of all the range references in the formula; each item is a range reference such as A1:A10
                        }
                        foreach (Match match in matchedRanges)
                        {
                            formula = formula.Replace(match.Value, "");
                            string ws_name = worksheet_name; // match.Value.Substring(0, match.Value.LastIndexOf("!")); // Get the name of the worksheet being referenced
                            string range_coordinates = match.Value.Substring(match.Value.LastIndexOf("!") + 1);  //match.Value.Replace(ws_name + "!", "");
                            string[] endCells = range_coordinates.Split(':');     //Split up each matched range into the start and end cells of the range
                            TreeNode range = null;
                            //Try to find the range in existing TreeNodes
                            foreach (TreeNode n in ranges)
                            {
                                if (n.getName().Replace("$", "") == endCells[0].Replace("$", "") + "_to_" + endCells[1].Replace("$", "") && n.getWorksheet() == ws_name)
                                {
                                    //MessageBox.Show("Found node:" + ws_name + "_" + endCells[0] + ":" + endCells[1]);
                                    range = n;
                                }
                                else
                                {
                                    continue;
                                }
                            }
                            //If it was not found, create it
                            if (range == null)
                            {
                                range = new TreeNode(endCells[0].Replace("$", "") + "_to_" + endCells[1].Replace("$", ""), ws_name, Globals.ThisAddIn.Application.ActiveWorkbook);
                                //MessageBox.Show("Created range node:" + ws_name + "_" + endCells[0] + ":" + endCells[1]);
                                ranges.Add(range);
                            }
                            chart_node.addParent(range);
                            range.addChild(chart_node);
                            //Add each cell contained in the range to the dependencies
                            foreach (Excel.Range cellInRange in Globals.ThisAddIn.Application.Worksheets[ws_index].Range[endCells[0], endCells[1]])
                            {
                                TreeNode input_cell = null;
                                //Find the node object for the current cell in the existing TreeNodes
                                //Check if this cell's coordinates are within the bounds of the used range, otherwise there will be an index out of bounds error
                                if (cellInRange.Column <= (cellInRange.Worksheet.UsedRange.Columns.Count + cellInRange.Worksheet.UsedRange.Column) && cellInRange.Row <= (cellInRange.Worksheet.UsedRange.Rows.Count + cellInRange.Worksheet.UsedRange.Row))
                                {
                                    //if a TreeNode exists for this cell already
                                    if (nodes_grid[cellInRange.Worksheet.Index - 1][cellInRange.Row - 1][cellInRange.Column - 1] != null)
                                    {
                                        input_cell = nodes_grid[cellInRange.Worksheet.Index - 1][cellInRange.Row - 1][cellInRange.Column - 1];
                                    }
                                }
                                //If it wasn't found, then it is blank, and we have to create a TreeNode for it
                                if (input_cell == null)
                                {
                                    input_cell = new TreeNode(cellInRange.Address, cellInRange.Worksheet.Name, Globals.ThisAddIn.Application.ActiveWorkbook);
                                    //Check if this cell's coordinates are within the bounds of the used range, otherwise there will be an index out of bounds error
                                    if (cellInRange.Column <= (cellInRange.Worksheet.UsedRange.Columns.Count + cellInRange.Worksheet.UsedRange.Column) && cellInRange.Row <= (cellInRange.Worksheet.UsedRange.Rows.Count + cellInRange.Worksheet.UsedRange.Row))
                                    {
                                        nodes_grid[cellInRange.Worksheet.Index - 1][cellInRange.Row - 1][cellInRange.Column - 1] = input_cell;
                                    }
                                }

                                //Update the dependencies
                                range.addParent(input_cell);
                                input_cell.addChild(range);
                            }
                        }

                        // Now we look for references of the kind 'worksheet_name'!A1 (with quotation marks)
                        if (toggle_compile_regex.Checked)
                        {
                            //Regex regex = new Regex(@"('" + worksheet_name + @"'!\$?[A-Z]+\$?[1-9]\d*:\$?[A-Z]+\$?[1-9]\d*)", RegexOptions.Compiled);
                            //Regex regex = regex_array[ws_index - 1];
                            matchedCells = regex_array[4*(ws_index - 1) + 2].Matches(formula);
                        }
                        else
                        {
                            matchedCells = Regex.Matches(formula, @"('" + worksheet_name + @"'!\$?[A-Z]+\$?[1-9]\d*)"); //matchedCells is a collection of all the references in the formula to cells in the specific worksheet, where the reference has the form 'worksheet_name'!A1
                        }
                        foreach (Match match in matchedCells)
                        {
                            formula = formula.Replace(match.Value, "");
                            string ws_name = worksheet_name; // match.Value.Substring(1, match.Value.LastIndexOf("!") - 2); // Get the name of the worksheet being referenced
                            string cell_coordinates = match.Value.Substring(match.Value.LastIndexOf("!") + 1);
                            //Get the actual cell that is being referenced
                            Excel.Range input = null;
                            foreach (Excel.Worksheet ws in Globals.ThisAddIn.Application.Worksheets)
                            {
                                //Find the worksheet object that the match belongs to
                                if (ws.Name == ws_name)
                                {
                                    input = ws.get_Range(cell_coordinates);
                                }
                            }
                            TreeNode input_cell = null;
                            //Find the node object for the current cell in the existing TreeNodes
                            //Check if this cell's coordinates are within the bounds of the used range of its spreadsheet, otherwise there will be an index out of bounds error
                            if (input.Column <= (input.Worksheet.UsedRange.Columns.Count + input.Worksheet.UsedRange.Column) && input.Row <= (input.Worksheet.UsedRange.Rows.Count + input.Worksheet.UsedRange.Row))
                            {
                                //if a TreeNode exists for this cell already, use it
                                if (nodes_grid[input.Worksheet.Index - 1][input.Row - 1][input.Column - 1] != null)
                                {
                                    input_cell = nodes_grid[input.Worksheet.Index - 1][input.Row - 1][input.Column - 1];
                                }
                            }
                            //If it wasn't found, then it is blank, and we have to create a TreeNode for it
                            if (input_cell == null)
                            {
                                input_cell = new TreeNode(cell_coordinates.Replace("$", ""), ws_name, Globals.ThisAddIn.Application.ActiveWorkbook);
                                //Check if this cell's coordinates are within the bounds of the used range, otherwise there will be an index out of bounds error
                                if (input.Column <= (input.Worksheet.UsedRange.Columns.Count + input.Worksheet.UsedRange.Column) && input.Row <= (input.Worksheet.UsedRange.Rows.Count + input.Worksheet.UsedRange.Row))
                                {
                                    nodes_grid[input.Worksheet.Index - 1][input.Row - 1][input.Column - 1] = input_cell;
                                }
                            }

                            //Update the dependencies
                            chart_node.addParent(input_cell);
                            input_cell.addChild(chart_node);
                        }

                        //Lastly we look for references of the kind worksheet_name!A1 (without quotation marks)
                        if (toggle_compile_regex.Checked)
                        {
                            //Regex regex = new Regex(@"(" + worksheet_name + @"!\$?[A-Z]+\$?[1-9]\d*)", RegexOptions.Compiled);
                            //Regex regex = regex_array[ws_index + 3 - 1];
                            matchedCells = regex_array[4*(ws_index - 1) + 3].Matches(formula);
                        }
                        else
                        {
                            matchedCells = Regex.Matches(formula, @"(" + worksheet_name + @"!\$?[A-Z]+\$?[1-9]\d*)");
                        }
                        foreach (Match match in matchedCells)
                        {
                            formula = formula.Replace(match.Value, "");
                            string ws_name = worksheet_name; // match.Value.Substring(0, match.Value.LastIndexOf("!")); // Get the name of the worksheet being referenced
                            string cell_coordinates = match.Value.Substring(match.Value.LastIndexOf("!") + 1);
                            //Get the actual cell that is being referenced
                            Excel.Range input = null;
                            foreach (Excel.Worksheet ws in Globals.ThisAddIn.Application.Worksheets)
                            {
                                //Find the worksheet object that the match belongs to
                                if (ws.Name == ws_name)
                                {
                                    input = ws.get_Range(cell_coordinates);
                                }
                            }
                            TreeNode input_cell = null;
                            //Find the node object for the current cell in the existing TreeNodes
                            //Check if this cell's coordinates are within the bounds of the used range of its spreadsheet, otherwise there will be an index out of bounds error
                            if (input.Column <= (input.Worksheet.UsedRange.Columns.Count + input.Worksheet.UsedRange.Column) && input.Row <= (input.Worksheet.UsedRange.Rows.Count + input.Worksheet.UsedRange.Row))
                            {
                                //if a TreeNode exists for this cell already, use it
                                if (nodes_grid[input.Worksheet.Index - 1][input.Row - 1][input.Column - 1] != null)
                                {
                                    input_cell = nodes_grid[input.Worksheet.Index - 1][input.Row - 1][input.Column - 1];
                                }
                            }
                            //If it wasn't found, then it is blank, and we have to create a TreeNode for it
                            if (input_cell == null)
                            {
                                input_cell = new TreeNode(cell_coordinates.Replace("$", ""), ws_name, Globals.ThisAddIn.Application.ActiveWorkbook);
                                //Check if this cell's coordinates are within the bounds of the used range, otherwise there will be an index out of bounds error
                                if (input.Column <= (input.Worksheet.UsedRange.Columns.Count + input.Worksheet.UsedRange.Column) && input.Row <= (input.Worksheet.UsedRange.Rows.Count + input.Worksheet.UsedRange.Row))
                                {
                                    nodes_grid[input.Worksheet.Index - 1][input.Row - 1][input.Column - 1] = input_cell;
                                }
                            }

                            //Update the dependencies
                            chart_node.addParent(input_cell);
                            input_cell.addChild(chart_node);
                        }
                        ws_index++;
                    }
                    //Find any references to named ranges
                    //TODO -- this should probably be done in a better way - with a regular expression that will catch things like this:
                    //"+range_name", "-range_name", "*range_name", etc., because right now a range name may be part of the name of a 
                    //formula that is used. For instance a range could be named "s", and if the formula has the "sum" function in it, we will 
                    //falsely detect a reference to "s". This does not affect the correctness of the algorithm, because all we care about 
                    //from the dependence graph is identifying which cells are outputs, and identifying user-defined ranges
                    //and this type of error will not affect either one
                    foreach (Excel.Name named_range in Globals.ThisAddIn.Application.Names)
                    {
                        if (formula.Contains(named_range.Name))
                        {
                            formula = formula.Replace(named_range.Name, "");
                        }
                        else
                        {
                            continue;
                        }

                        string[] endCells = named_range.RefersToRange.Address.Split(':');     //Split up each matched range into the start and end cells of the range
                        TreeNode range = null;
                        //Try to find the range in existing TreeNodes
                        foreach (TreeNode n in ranges)
                        {
                            if (n.getName().Replace("$", "") == endCells[0].Replace("$", "") + "_to_" + endCells[1].Replace("$", "") && n.getWorksheet() == named_range.RefersToRange.Worksheet.Name)
                            {
                                range = n;
                            }
                            else
                            {
                                continue;
                            }
                        }
                        //If it does not exist, create it
                        if (range == null)
                        {
                            range = new TreeNode(endCells[0].Replace("$", "") + "_to_" + endCells[1].Replace("$", ""), named_range.RefersToRange.Worksheet.Name, Globals.ThisAddIn.Application.ActiveWorkbook);
                            ranges.Add(range);
                        }
                        //Update dependencies
                        chart_node.addParent(range);
                        range.addChild(chart_node);
                        //Add each cell contained in the range to the dependencies
                        foreach (Excel.Range cellInRange in named_range.RefersToRange.Worksheet.Range[endCells[0], endCells[1]])
                        {
                            TreeNode input_cell = null;
                            //Find the node object for the current cell in the existing TreeNodes
                            //Check if this cell's coordinates are within the bounds of the used range, otherwise there will be an index out of bounds error
                            if (cellInRange.Column <= (cellInRange.Worksheet.UsedRange.Columns.Count + cellInRange.Worksheet.UsedRange.Column) && cellInRange.Row <= (cellInRange.Worksheet.UsedRange.Rows.Count + cellInRange.Worksheet.UsedRange.Row))
                            {
                                //if a TreeNode exists for this cell already, use it
                                if (nodes_grid[cellInRange.Worksheet.Index - 1][cellInRange.Row - 1][cellInRange.Column - 1] != null)
                                {
                                    input_cell = nodes_grid[cellInRange.Worksheet.Index - 1][cellInRange.Row - 1][cellInRange.Column - 1];
                                }
                            }
                            //If it wasn't found, then it is blank, and we have to create a TreeNode for it
                            if (input_cell == null)
                            {
                                input_cell = new TreeNode(cellInRange.Address, cellInRange.Worksheet.Name, Globals.ThisAddIn.Application.ActiveWorkbook);
                                //Check if this cell's coordinates are within the bounds of the used range, otherwise there will be an index out of bounds error
                                if (cellInRange.Column <= (cellInRange.Worksheet.UsedRange.Columns.Count + cellInRange.Worksheet.UsedRange.Column) && cellInRange.Row <= (cellInRange.Worksheet.UsedRange.Rows.Count + cellInRange.Worksheet.UsedRange.Row))
                                {
                                    nodes_grid[cellInRange.Worksheet.Index - 1][cellInRange.Row - 1][cellInRange.Column - 1] = input_cell;
                                }
                            }

                            //Update the dependencies
                            range.addParent(input_cell);
                            input_cell.addChild(range);
                        }
                    }
                    //In this case every reference to cells or ranges must explicitly state their worksheet, so no additional analysis is necessary 
                }
            }
            foreach (Excel.Worksheet worksheet in Globals.ThisAddIn.Application.Worksheets)
            {
                foreach (Excel.ChartObject chart in (Excel.ChartObjects)worksheet.ChartObjects(Type.Missing))
                {
                    //TODO The naming convention for TreeNode charts is kind of a hack; could fail if two charts have the same names when white spaces are removed - maybe add a random hash at the end
                    TreeNode chart_node = new TreeNode(chart.Name, worksheet.Name, Globals.ThisAddIn.Application.ActiveWorkbook);
                    chart_node.setChart(true);
                    nodes.Add(chart_node);
                    foreach (Excel.Series series in (Excel.SeriesCollection)chart.Chart.SeriesCollection(Type.Missing))
                    {
                        string formula = "";
                        try
                        {
                            formula = series.Formula;  //The formula contained in the cell
                        }
                        catch
                        {

                        }
                        //find_references(chart_node, formula);
                        
                        MatchCollection matchedRanges = null;
                        MatchCollection matchedCells = null;
                        int ws_index = 1;
                        foreach (string s in worksheet_names)
                        {
                            string worksheet_name = s.Replace("+", @"\+").Replace("^", @"\^").Replace("$", @"\$").Replace(".", @"\."); //Escape certain characters in the regular expression
                            //First look for range references of the form 'worksheet_name'!A1:A10 in the formula (with quotation marks around the name)
                            matchedRanges = Regex.Matches(formula, @"('" + worksheet_name + @"'!\$?[A-Z]+\$?[1-9]\d*:\$?[A-Z]+\$?[1-9]\d*)");  //A collection of all the references in the formula to ranges in the particular worksheet; each item is a range reference of the form 'worksheet_name'!A1:A10
                            foreach (Match match in matchedRanges)
                            {
                                formula = formula.Replace(match.Value, "");
                                string ws_name = worksheet_name; // match.Value.Substring(1, match.Value.LastIndexOf("!") - 2); // Get the name of the worksheet being referenced
                                string range_coordinates = match.Value.Substring(match.Value.LastIndexOf("!") + 1); //match.Value.Replace("'" + ws_name + "'!", "");
                                string[] endCells = range_coordinates.Split(':');     //Split up each matched range into the start and end cells of the range
                                TreeNode range = null;
                                //Try to find the range in existing TreeNodes
                                foreach (TreeNode n in nodes)
                                {
                                    if (n.getName().Replace("$", "") == endCells[0].Replace("$", "") + "_to_" + endCells[1].Replace("$", "") && n.getWorksheet() == ws_name)
                                    {
                                        //MessageBox.Show("Found node:" + ws_name + "_" + endCells[0] + ":" + endCells[1]);
                                        range = n;
                                    }
                                    else
                                    {
                                        continue;
                                    }
                                }
                                //If it does not exist, create it
                                if (range == null)
                                {
                                    range = new TreeNode(endCells[0].Replace("$", "") + "_to_" + endCells[1].Replace("$", ""), ws_name, Globals.ThisAddIn.Application.ActiveWorkbook);
                                    //MessageBox.Show("Created range node:" + ws_name + "_" + endCells[0] + ":" + endCells[1]);
                                    nodes.Add(range);
                                }
                                chart_node.addParent(range);
                                range.addChild(chart_node);
                                //Add each cell contained in the range to the dependencies
                                foreach (Excel.Range cellInRange in Globals.ThisAddIn.Application.Worksheets[ws_index].Range[endCells[0], endCells[1]])
                                {
                                    TreeNode input_cell = null;
                                    //Find the node object for the current cell in the list of TreeNodes
                                    foreach (TreeNode node in nodes)
                                    {
                                        if (node.getName().Replace("$", "") == cellInRange.Address.Replace("$", "") && node.getWorksheet() == cellInRange.Worksheet.Name)
                                        {
                                            input_cell = node;
                                        }
                                        else
                                            continue;
                                    }
                                    //If it wasn't found, then it is blank, and we have to create a TreeNode for it
                                    if (input_cell == null)
                                    {
                                        input_cell = new TreeNode(cellInRange.Address, cellInRange.Worksheet.Name, Globals.ThisAddIn.Application.ActiveWorkbook);
                                        nodes.Add(input_cell);
                                    }

                                    //Update the dependencies
                                    range.addParent(input_cell);
                                    input_cell.addChild(range);
                                }
                            }

                            //Next look for range references of the form worksheet_name!A1:A10 in the formula (no quotation marks around the name)
                            matchedRanges = Regex.Matches(formula, @"(" + worksheet_name + @"!\$?[A-Z]+\$?[1-9]\d*:\$?[A-Z]+\$?[1-9]\d*)");  //A collection of all the range references in the formula; each item is a range reference such as A1:A10
                            foreach (Match match in matchedRanges)
                            {
                                formula = formula.Replace(match.Value, "");
                                string ws_name = worksheet_name; // match.Value.Substring(0, match.Value.LastIndexOf("!")); // Get the name of the worksheet being referenced
                                string range_coordinates = match.Value.Substring(match.Value.LastIndexOf("!") + 1);  //match.Value.Replace(ws_name + "!", "");
                                string[] endCells = range_coordinates.Split(':');     //Split up each matched range into the start and end cells of the range
                                TreeNode range = null;
                                //Try to find the range in existing TreeNodes
                                foreach (TreeNode n in nodes)
                                {
                                    if (n.getName().Replace("$", "") == endCells[0].Replace("$", "") + "_to_" + endCells[1].Replace("$", "") && n.getWorksheet() == ws_name)
                                    {
                                        //MessageBox.Show("Found node:" + ws_name + "_" + endCells[0] + ":" + endCells[1]);
                                        range = n;
                                    }
                                    else
                                    {
                                        continue;
                                    }
                                }
                                //If it does not exist, create it
                                if (range == null)
                                {
                                    range = new TreeNode(endCells[0].Replace("$", "") + "_to_" + endCells[1].Replace("$", ""), ws_name, Globals.ThisAddIn.Application.ActiveWorkbook);
                                    //MessageBox.Show("Created node:" + ws_name + "_" + endCells[0] + ":" + endCells[1]);
                                    nodes.Add(range);
                                }

                                //Update the dependencies
                                chart_node.addParent(range);
                                range.addChild(chart_node);
                                //Add each cell contained in the range to the dependencies
                                foreach (Excel.Range cellInRange in Globals.ThisAddIn.Application.Worksheets[ws_index].Range[endCells[0], endCells[1]])
                                {
                                    TreeNode input_cell = null;
                                    //Find the node object for the current cell in the list of TreeNodes
                                    foreach (TreeNode node in nodes)
                                    {
                                        if (node.getName().Replace("$", "") == cellInRange.Address.Replace("$", "") && node.getWorksheet() == cellInRange.Worksheet.Name)
                                        {
                                            input_cell = node;
                                        }
                                        else
                                            continue;
                                    }
                                    //If it wasn't found, then it is blank, and we have to create a TreeNode for it
                                    if (input_cell == null)
                                    {
                                        input_cell = new TreeNode(cellInRange.Address, cellInRange.Worksheet.Name, Globals.ThisAddIn.Application.ActiveWorkbook);
                                        nodes.Add(input_cell);
                                    }

                                    //Update the dependencies
                                    range.addParent(input_cell);
                                    input_cell.addChild(range);
                                }
                            }

                            // Now we look for references of the kind 'worksheet_name'!A1 (with quotation marks)
                            matchedCells = Regex.Matches(formula, @"('" + worksheet_name + @"'!\$?[A-Z]+\$?[1-9]\d*)"); //matchedCells is a collection of all the references in the formula to cells in the specific worksheet, where the reference has the form 'worksheet_name'!A1
                            foreach (Match match in matchedCells)
                            {
                                formula = formula.Replace(match.Value, "");
                                string ws_name = worksheet_name; // match.Value.Substring(1, match.Value.LastIndexOf("!") - 2); // Get the name of the worksheet being referenced
                                string cell_coordinates = match.Value.Substring(match.Value.LastIndexOf("!") + 1);

                                TreeNode input_cell = null;
                                //Find the node object for the current cell in the list of TreeNodes
                                foreach (TreeNode node in nodes)
                                {
                                    if (node.getName().Replace("$", "") == cell_coordinates.Replace("$", "") && node.getWorksheet() == ws_name)
                                    {
                                        input_cell = node;
                                    }
                                    else
                                    {
                                        continue;
                                    }
                                }
                                //If it wasn't found, then it is blank, and we have to create a TreeNode for it
                                if (input_cell == null)
                                {
                                    input_cell = new TreeNode(cell_coordinates.Replace("$", ""), ws_name, Globals.ThisAddIn.Application.ActiveWorkbook);
                                    nodes.Add(input_cell);
                                }

                                //Update the dependencies
                                chart_node.addParent(input_cell);
                                input_cell.addChild(chart_node);
                            }

                            //Lastly we look for references of the kind worksheet_name!A1 (without quotation marks)
                            matchedCells = Regex.Matches(formula, @"(" + worksheet_name + @"!\$?[A-Z]+\$?[1-9]\d*)");
                            foreach (Match match in matchedCells)
                            {
                                formula = formula.Replace(match.Value, "");
                                string ws_name = worksheet_name; // match.Value.Substring(0, match.Value.LastIndexOf("!")); // Get the name of the worksheet being referenced
                                string cell_coordinates = match.Value.Substring(match.Value.LastIndexOf("!") + 1);

                                TreeNode input_cell = null;
                                //Find the node object for the current cell in the list of TreeNodes
                                foreach (TreeNode node in nodes)
                                {
                                    if (node.getName().Replace("$", "") == cell_coordinates.Replace("$", "") && node.getWorksheet() == ws_name)
                                    {
                                        input_cell = node;
                                    }
                                    else
                                    {
                                        continue;
                                    }
                                }
                                //If it wasn't found, then it is blank, and we have to create a TreeNode for it
                                if (input_cell == null)
                                {
                                    input_cell = new TreeNode(cell_coordinates.Replace("$", ""), ws_name, Globals.ThisAddIn.Application.ActiveWorkbook);
                                    nodes.Add(input_cell);
                                }

                                //Update the dependencies
                                chart_node.addParent(input_cell);
                                input_cell.addChild(chart_node);
                            }
                            ws_index++;
                        }
                        string patternRange = @"(\$?[A-Z]+\$?[1-9]\d*:\$?[A-Z]+\$?[1-9]\d*)";  //Regex for matching range references in formulas such as A1:A10, or $A$1:$A$10 etc.
                        string patternCell = @"(\$?[A-Z]+\$?[1-9]\d*)";        //Regex for matching single cell references such as A1 or $A$1, etc. 

                        //First look for range references in the formula
                        matchedRanges = Regex.Matches(formula, patternRange);  //A collection of all the range references in the formula; each item is a range reference such as A1:A10
                        List<Excel.Range> rangeList = new List<Excel.Range>();
                        foreach (Match match in matchedRanges)
                        {
                            formula = formula.Replace(match.Value, "");
                            string[] endCells = match.Value.Split(':');     //Split up each matched range into the start and end cells of the range
                            TreeNode range = null;
                            //Try to find the range in existing TreeNodes
                            foreach (TreeNode node in nodes)
                            {
                                if (node.getName() == endCells[0].Replace("$", "") + "_to_" + endCells[1].Replace("$", "") && node.getWorksheet() == worksheet.Name)
                                {
                                    range = node;
                                }
                                else
                                {
                                    continue;
                                }
                            }
                            //If it does not exist, create it
                            if (range == null)
                            {
                                range = new TreeNode(endCells[0].Replace("$", "") + "_to_" + endCells[1].Replace("$", ""), worksheet.Name, Globals.ThisAddIn.Application.ActiveWorkbook);
                                nodes.Add(range);
                            }

                            //Update the dependencies
                            chart_node.addParent(range);
                            range.addChild(chart_node);
                            //Add each cell contained in the range to the dependencies
                            foreach (Excel.Range cellInRange in worksheet.Range[endCells[0], endCells[1]])
                            {
                                TreeNode input_cell = null;
                                //Find the node object for the current cell in the list of TreeNodes
                                foreach (TreeNode node in nodes)
                                {
                                    if (node.getName().Replace("$", "") == cellInRange.Address.Replace("$", "") && node.getWorksheet() == worksheet.Name)
                                    {
                                        input_cell = node;
                                    }
                                    else
                                        continue;
                                }
                                //If it wasn't found, then it is blank, and we have to create a TreeNode for it
                                if (input_cell == null)
                                {
                                    input_cell = new TreeNode(cellInRange.Address, cellInRange.Worksheet.Name, Globals.ThisAddIn.Application.ActiveWorkbook);
                                    nodes.Add(input_cell);
                                }

                                //Update the dependencies
                                range.addParent(input_cell);
                                input_cell.addChild(range);
                            }
                        }

                        //Find any references to named ranges
                        //TODO -- this should probably be done in a better way - with a regular expression that will catch things like this:
                        //"+range_name", "-range_name", "*range_name", etc., because right now a range name may be part of the name of a 
                        //formula that is used. For instance a range could be named "s", and if the formula has the "sum" function in it, we will 
                        //falsely detect a reference to "s". This does not affect the correctness of the algorithm, because all we care about 
                        //from the dependence graph is identifying which cells are outputs, and identifying user-defined ranges
                        //and this type of error will not affect either one
                        foreach (Excel.Name named_range in Globals.ThisAddIn.Application.Names)
                        {
                            if (formula.Contains(named_range.Name))
                            {
                                formula = formula.Replace(named_range.Name, "");
                            }
                            else
                            {
                                continue;
                            }

                            string[] endCells = named_range.RefersToRange.Address.Split(':');     //Split up each matched range into the start and end cells of the range
                            TreeNode range = null;
                            //Try to find the range in existing TreeNodes
                            foreach (TreeNode n in ranges)
                            {
                                if (n.getName().Replace("$", "") == endCells[0].Replace("$", "") + "_to_" + endCells[1].Replace("$", "") && n.getWorksheet() == named_range.RefersToRange.Worksheet.Name)
                                {
                                    range = n;
                                }
                                else
                                {
                                    continue;
                                }
                            }
                            //If it does not exist, create it
                            if (range == null)
                            {
                                //MessageBox.Show("Created range node:" + c.Worksheet.Name + "_" + endCells[0] + ":" + endCells[1]);
                                range = new TreeNode(endCells[0].Replace("$", "") + "_to_" + endCells[1].Replace("$", ""), named_range.RefersToRange.Worksheet.Name, Globals.ThisAddIn.Application.ActiveWorkbook);
                                ranges.Add(range);
                            }
                            //Update dependencies
                            chart_node.addParent(range);
                            range.addChild(chart_node);
                            //Add each cell contained in the range to the dependencies
                            foreach (Excel.Range cellInRange in named_range.RefersToRange.Worksheet.Range[endCells[0], endCells[1]])
                            {
                                TreeNode input_cell = null;
                                //Find the node object for the current cell in the existing TreeNodes
                                //Check if this cell's coordinates are within the bounds of the used range, otherwise there will be an index out of bounds error
                                if (cellInRange.Column <= (cellInRange.Worksheet.UsedRange.Columns.Count + cellInRange.Worksheet.UsedRange.Column) && cellInRange.Row <= (cellInRange.Worksheet.UsedRange.Rows.Count + cellInRange.Worksheet.UsedRange.Row))
                                {
                                    //if a TreeNode exists for this cell already, use it
                                    if (nodes_grid[cellInRange.Worksheet.Index - 1][cellInRange.Row - 1][cellInRange.Column - 1] != null)
                                    {
                                        input_cell = nodes_grid[cellInRange.Worksheet.Index - 1][cellInRange.Row - 1][cellInRange.Column - 1];
                                    }
                                }
                                //If it wasn't found, then it is blank, and we have to create a TreeNode for it
                                if (input_cell == null)
                                {
                                    input_cell = new TreeNode(cellInRange.Address, cellInRange.Worksheet.Name, Globals.ThisAddIn.Application.ActiveWorkbook);
                                    //Check if this cell's coordinates are within the bounds of the used range, otherwise there will be an index out of bounds error
                                    if (cellInRange.Column <= (cellInRange.Worksheet.UsedRange.Columns.Count + cellInRange.Worksheet.UsedRange.Column) && cellInRange.Row <= (cellInRange.Worksheet.UsedRange.Rows.Count + cellInRange.Worksheet.UsedRange.Row))
                                    {
                                        nodes_grid[cellInRange.Worksheet.Index - 1][cellInRange.Row - 1][cellInRange.Column - 1] = input_cell;
                                    }
                                }

                                //Update the dependencies
                                range.addParent(input_cell);
                                input_cell.addChild(range);
                            }
                        }

                        matchedCells = Regex.Matches(formula, patternCell);  //matchedCells is a collection of all the cells that are referenced by the formula
                        foreach (Match m in matchedCells)
                        {
                            TreeNode input_cell = null;
                            //Find the node object for the current cell in the list of TreeNodes
                            foreach (TreeNode node in nodes)
                            {
                                if (node.getName().Replace("$", "") == m.Value.Replace("$", "") && node.getWorksheet() == worksheet.Name)
                                {
                                    input_cell = node;
                                }
                                else
                                {
                                    continue;
                                }
                            }
                            //If it wasn't found, then it is blank, and we have to create a TreeNode for it
                            if (input_cell == null)
                            {
                                input_cell = new TreeNode(m.Value, worksheet.Name, Globals.ThisAddIn.Application.ActiveWorkbook);
                                nodes.Add(input_cell);
                            }
                            //Update the dependencies
                            chart_node.addParent(input_cell);
                            input_cell.addChild(chart_node);
                        }

                    }
                }
            }
            
            //TODO -- we are not able to capture ranges that are identified in stored procedures or macros, just ones referenced in formulas
            //TODO -- Dealing with fuzzing of charts -- idea: any cell that feeds into a chart is essentially an output; the chart is just a visual representation (can charts operate on values before they are displayed? don't think so...)
            starting_outputs = new List<StartValue>(); //This will store the values of all the output nodes at the start of the procedure for swapping values (fuzzing)
            output_cells = new List<TreeNode>(); //This will store all the output nodes at the start of the fuzzing procedure

            //Store all the starting output values
            foreach (TreeNode[][] node_arr_arr in nodes_grid)
            {
                if (node_arr_arr != null)
                {
                    foreach (TreeNode[] node_arr in node_arr_arr)
                    {
                        foreach (TreeNode node in node_arr)
                        {
                            if (node != null)
                            {
                                if (!node.hasChildren() && node.hasParents()) //Nodes that do not feed into any other nodes are considered output, unless nothing feeds into them either. 
                                {
                                    output_cells.Add(node);
                                }
                            }
                        }
                    }
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
            //Tree building stopwatch
            TimeSpan tree_building_timespan = global_stopwatch.Elapsed;

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
            
            Globals.ThisAddIn.Application.ScreenUpdating = false;
            
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
                            //MessageBox.Show("output cells count = " + output_cells.Count);
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
                foreach (TreeNode[][] node_arr_arr in nodes_grid)
                {
                    if (node_arr_arr != null)
                    {
                        foreach (TreeNode[] node_arr in node_arr_arr)
                        {
                            foreach (TreeNode node in node_arr)
                            {
                                if (node != null)
                                {
                                    if (!node.hasParents())
                                    {
                                        node.setWeight(1.0);  //Set the weight of all input nodes to 1.0 to start
                                        //Now we propagate it's weight to all of it's children
                                        propagateWeightUp(node, 1.0, node);
                                        raw_input_cells_in_computation_count++;
                                    }
                                }
                            }
                        }
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
                    int swaps_per_range = 1; // 30;
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
                        if (swaps_per_range == 1)
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
                //                        //MessageBox.Show(worksheet.Name + ":R" + (row + 1) + "C" + (col + 1) + " is an outlier with respect to output " + (i + 1) + " with a z-score of " + impacts_grid[worksheet.Index - 1][row][col][i]);
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
                            //MessageBox.Show(worksheet.Name + ":R" + (row + 1) + "C" + (col + 1) + " is an outlier with respect to output " + (i + 1) + " with a z-score of " + impacts_grid[worksheet.Index - 1][row][col][i]);
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
            //MessageBox.Show("Done building dependence graph.\nTime elapsed: " + treeTime);
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
            nodes = new List<TreeNode>();        //This is a list holding all the TreeNodes in the Excel file

            ranges = new List<TreeNode>();        //This is a list holding all the ranges of TreeNodes in the Excel file
            charts = new List<TreeNode>();        //This is a list holding all the chart TreeNodes in the Excel file
            nodes_grid = new TreeNode[Globals.ThisAddIn.Application.Worksheets.Count + Globals.ThisAddIn.Application.Charts.Count][][];
            int index = 0;
            foreach (Excel.Worksheet worksheet in Globals.ThisAddIn.Application.Worksheets)
            {
                nodes_grid[worksheet.Index - 1] = new TreeNode[worksheet.UsedRange.Rows.Count + worksheet.UsedRange.Row][];
                for (int row = 0; row < (worksheet.UsedRange.Rows.Count + worksheet.UsedRange.Row); row++)
                {
                    nodes_grid[worksheet.Index - 1][row] = new TreeNode[worksheet.UsedRange.Columns.Count + worksheet.UsedRange.Column];
                    for (int col = 0; col < (worksheet.UsedRange.Columns.Count + worksheet.UsedRange.Column); col++)
                    {
                        nodes_grid[worksheet.Index - 1][row][col] = null;
                    }
                }
                index++;
            }
            
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
                    TreeNode n = new TreeNode(cell.Address, cell.Worksheet.Name, Globals.ThisAddIn.Application.ActiveWorkbook);  //Create a TreeNode for every cell with the name being the cell's address and set the node's worksheet appropriately
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

        //Action for the "Derivatives" button
        private void button2_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Worksheet activeWorksheet = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet;  //Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets[1]);
            Excel.Range selection = Globals.ThisAddIn.Application.Selection as Excel.Range;

            //If there is exactly one column in the selection
            if (selection.Columns.Count == 1)
            {
                foreach (Excel.Range cell in selection)
                {
                    Excel.Range cellUnder = cell.get_Offset(1, 0);
                    Excel.Range cellRight = cell.get_Offset(0, 1);
                    if (Globals.ThisAddIn.Application.Intersect(cellUnder, selection, Type.Missing,
                                    Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                    Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                    Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                    Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing) != null)
                    {
                        cellRight.Value = (cellUnder.Value - cell.Value);
                    }
                }
            }
            //If there are exactly two columns in the selection
            else if (selection.Columns.Count == 2)
            {
                int i = 0;
                String col_address = "";
                //This figures out the correct index column -- we take the leftmost to be the index column
                foreach (Excel.Range column in selection.Columns)
                {
                    i = i + 1;
                    if (i != 1)
                    {
                        continue;
                    }
                    col_address = column.Address;
                }
                //This loops through all the cells
                foreach (Excel.Range cell in selection)
                {
                    String cell_address = cell.Address;
                    //We have to parse the cell address to extract the coordinates; An example address is $B$9, but the oolumn may consist of
                    //Multiple letters such as $AA$94
                    string[] cell_coordinates = cell_address.Split('$'); //cell_coordinates is now as follows: [ -blank- , -column address-, -row address- ]
                    //We also have to parse row_address in a similar way; an example of row_address is $B$9:$H$9
                    string[] col_coordinates = col_address.Split('$', ':'); //col_coordinates is now as follows: [ -blank- , -column address 1-, -row address 1-,  -blank- , -column address 2-, -row address 2- ]
                    if (cell_coordinates[1] == col_coordinates[1])
                    {
                        Excel.Range cellUnder = cell.get_Offset(1, 0);
                        Excel.Range cellRight = cell.get_Offset(0, 1);
                        Excel.Range cellRightRight = cell.get_Offset(0, 2);
                        Excel.Range cellRightUnder = cell.get_Offset(1, 1);
                        if (Globals.ThisAddIn.Application.Intersect(cellUnder, selection, Type.Missing,
                                        Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                        Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                        Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                        Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing) != null)
                        {
                            if (cellUnder.Value - cell.Value != 0)
                            {
                                cellRightRight.Value = ((cellRightUnder.Value - cellRight.Value) / (cellUnder.Value - cell.Value));
                            }
                        }
                    }
                }
            }
            //If there is exactly one row in the selection
            else if (selection.Rows.Count == 1)
            {
                foreach (Excel.Range cell in selection)
                {
                    Excel.Range cellUnder = cell.get_Offset(1, 0);
                    Excel.Range cellRight = cell.get_Offset(0, 1);
                    if (Globals.ThisAddIn.Application.Intersect(cellRight, selection, Type.Missing,
                                    Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                    Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                    Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                    Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing) != null)
                    {
                        cellUnder.Value = (cellRight.Value - cell.Value);
                    }
                }
            }
            //If there are exactly two rows in the selection
            else if (selection.Rows.Count == 2)
            {
                int i = 0;
                String row_address = "";
                //This figures out the correct index row -- the top row is used as the index row
                foreach (Excel.Range row in selection.Rows)
                {
                    i = i + 1;
                    if (i != 1)
                    {
                        continue;
                    }
                    row_address = row.Address;
                }
                //This loops through all the cells
                foreach (Excel.Range cell in selection)
                {
                    String cell_address = cell.Address;
                    //We have to parse the cell address to extract the coordinates; An example address is $B$9, but the oolumn may consist of
                    //Multiple letters such as $AA$94
                    string[] cell_coordinates = cell_address.Split('$'); //cell_coordinates is now as follows: [ -blank- , -column address-, -row address- ]
                    //We also have to parse row_address in a similar way; an example of row_address is $B$9:$H$9
                    string[] row_coordinates = row_address.Split('$', ':'); //row_coordinates is now as follows: [ -blank- , -column address 1-, -row address 1-,  -blank- , -column address 2-, -row address 2- ]
                    if (cell_coordinates[2] == row_coordinates[2])
                    {
                        Excel.Range cellUnder = cell.get_Offset(1, 0);
                        Excel.Range cellRight = cell.get_Offset(0, 1);
                        Excel.Range cellUnderUnder = cell.get_Offset(2, 0);
                        Excel.Range cellRightUnder = cell.get_Offset(1, 1);
                        if (Globals.ThisAddIn.Application.Intersect(cellRight, selection, Type.Missing,
                                        Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                        Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                        Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                        Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing) != null)
                        {
                            if (cellRight.Value - cell.Value != 0)
                            {
                                cellUnderUnder.Value = ((cellRightUnder.Value - cellUnder.Value) / (cellRight.Value - cell.Value));
                                cellUnderUnder.Interior.Color = System.Drawing.Color.AliceBlue;
                            }
                        }
                    }
                }
            }
        }

        /*
         * * * * * * * * STATISTICAL THINGS BEGIN HERE ;) * * * * * * * * *
         */

        //Dictionary stores the initial colors of all the cells so they can be restored by pressing the "Clear" button
        private Dictionary<Excel.Range, System.Drawing.Color> startColors = new Dictionary<Excel.Range, System.Drawing.Color>();
        
        private void button3_Click(object sender, RibbonControlEventArgs e)
        {
            //Performs the Anderson-Darling test for normality
            //Reject if AD > CV = 0.752 / (1 + 0.75/n + 2.25/(n^2) )
            //AD = SUM[i=1 to n] (1-2i)/n * {ln(F0[z_i]) + ln(1-F0[Z_(n+1-i)]) } - n
            // get user selection
            Excel.Range selection = Globals.ThisAddIn.Application.Selection as Excel.Range;

            // assume that the cells are normally distributed
            Stats.NormalAD normalAD = new Stats.NormalAD(selection);
        }

        Dictionary<Excel.Range, System.Drawing.Color> outliers;
        Boolean first_run = true;  // We only want to store the starting colors once, so this boolean is used for checking that
        private void button4_Click(object sender, RibbonControlEventArgs e)
        {
            if (first_run == true)      //if this is the first time running the test, store the starting colors of all cells
            {
                foreach (Excel.Range cell in ((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).UsedRange)
                {
                    startColors.Add(cell, System.Drawing.ColorTranslator.FromOle((int)cell.Interior.Color));
                }
                first_run = false;      // Update the boolean value to remember that we have run the test once already
            }

            // get user selection
            Excel.Range selection = Globals.ThisAddIn.Application.Selection as Excel.Range;

            // assume that the cells are normally distributed
            Stats.NormalDistribution norm_d = new Stats.NormalDistribution(selection);

            // find outliers
            outliers = norm_d.PeirceOutliers();

            // color the cells pink
            Stats.Utilities.ColorCellListByName(outliers, "pink");
        }

        private void button5_Click(object sender, RibbonControlEventArgs e)
        {
            //TODO need to revise the "Clear" button functionality, because if it is pressed after the "Analyze worksheet" button and cells are already colored, pressing "Clear" gives an error
            //Restore original color to cells flagged as outliers
            Stats.Utilities.RestoreColor(startColors);   
        }

        private void button6_Click(object sender, RibbonControlEventArgs e)
        {
            // get user selection
            Excel.Range selection = Globals.ThisAddIn.Application.Selection as Excel.Range;

            // assume that the cells are normally distributed
            Stats.NormalKS normalKS = new Stats.NormalKS(selection);
        }

        //Button for testing random code :)
        private void button7_Click(object sender, RibbonControlEventArgs e)
        {
            MessageBox.Show(Globals.ThisAddIn.Application.ActiveWorkbook.Path + "");
            MessageBox.Show(Globals.ThisAddIn.Application.Workbooks[1] + "");
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
                        //MessageBox.Show(node.getName() + " " + node.getOriginalColor());
                        //node.getWorksheetObject().get_Range(node.getName()).Interior.ColorIndex = 0;
                        //node.getWorksheetObject().get_Range(node.getName()).Interior.ColorIndex = node.getOriginalColor();
                        if (!(originalColorNodes[i].getOriginalColor() + "").Equals("Color [White]"))
                        {
                            originalColorNodes[i].getWorksheetObject().get_Range(originalColorNodes[i].getName()).Interior.Color = originalColorNodes[i].getOriginalColor();
                        }
                        else
                        {
                            originalColorNodes[i].getWorksheetObject().get_Range(originalColorNodes[i].getName()).Interior.ColorIndex = -4142;
                        }
                    }
                    originalColorNodes.RemoveAt(i);
                    i--;
                }
            }
        }

        private void peirce_button_Click(object sender, RibbonControlEventArgs e)
        {
            //run_peirce(Globals.ThisAddIn.Application.Selection as Excel.Range);
            //get_peirce_cutoff((Globals.ThisAddIn.Application.Selection as Excel.Range).Cells.Count);
            //MessageBox.Show("" + (Globals.ThisAddIn.Application.Selection as Excel.Range).Cells.Count);
            /**
            Excel.Range range = Globals.ThisAddIn.Application.Selection as Excel.Range;
            int m = 1;
            int k = 1;
            int N = range.Rows.Count;
            double precision1 = Math.Pow(10.0, -10.0);
            double precision2 = Math.Pow(10.0, -16.0);
            MessageBox.Show("" + N);
            if (N - m - k <= 0)
            {
                MessageBox.Show("Cutoff undefined.");
            }

            double LnQN = k * Math.Log(k, Math.E) + (N - k) * (Math.Log(N - k, Math.E)) - N * Math.Log(N, Math.E);
            double x = 1;
            double oldx;
            do
            {
                x = Math.Min(x, Math.Sqrt((N - m) / k) - precision1);

                //R1(x) and R2(x)
                double R1 = Math.Exp((x * x - 1) / 2) * DataDebug.Stats.Utilities.erfc(x / Math.Sqrt(2));
                //MessageBox.Show("Argument: " + x / Math.Sqrt(2)
                    //+ "\nERFC(Argument) = " + DataDebug.Stats.Utilities.erfc(x/Math.Sqrt(2)));
                double R2 = Math.Exp( (LnQN - 0.5 * (N - k) * Math.Log((N - m - k * x * x) / (N - m - k), Math.E)) / k);

                //R1'(x) and R2'(x)
                double R1d = x * R1 - Math.Sqrt(2 / Math.PI / Math.Exp(1));
                double R2d = x * (N - k) / (N - m - k * x * x) * R2;

                oldx = x;
                x = oldx - (R1 - R2) / (R1d - R2d);
                //MessageBox.Show("x = " + x);
            } while (Math.Abs(x - oldx) > N * 2 * precision2);
            MessageBox.Show("Done: x = " + x);
             **/
        }

        private double get_peirce_cutoff(int N, int m, int k)
        {
            double precision1 = Math.Pow(10.0, -10.0);
            double precision2 = Math.Pow(10.0, -16.0);
            if (N - m - k <= 0)
            {
                return 0; 
            }

            double LnQN = k * Math.Log(k, Math.E) + (N - k) * (Math.Log(N - k, Math.E)) - N * Math.Log(N, Math.E);
            double x = 1;
            double oldx;
            int counter = 0; //keep track of how many iterations of newton's method have been done
            do
            {
                counter++;
                if (counter > 1000) {
                    MessageBox.Show("Newton's method is taking too long for N = " + N + ", k = " + k + ", m = " + m + ".");
                    if (k > 1)
                    {
                        //MessageBox.Show("Calculating approximate cutoff (average of adjacent cutoffs).");
                        x = (get_peirce_cutoff(N, m, k - 1) + get_peirce_cutoff(N, m, k + 1)) / 2;
                        return x;
                    }
                    else
                    {
                        return 0; 
                    }
                }

                x = Math.Min(x, Math.Sqrt((N - m) / k) - precision1);

                //R1(x) and R2(x)
                double R1 = Math.Exp((x * x - 1) / 2) * DataDebug.Stats.Utilities.erfc(x / Math.Sqrt(2));
                double R2 = Math.Exp((LnQN - 0.5 * (N - k) * Math.Log((N - m - k * x * x) / (N - m - k), Math.E)) / k);

                //R1'(x) and R2'(x)
                double R1d = x * R1 - Math.Sqrt(2 / Math.PI / Math.Exp(1));
                double R2d = x * (N - k) / (N - m - k * x * x) * R2;

                oldx = x;
                x = oldx - (R1 - R2) / (R1d - R2d);
            } while (Math.Abs(x - oldx) > N * 2 * precision2);
            return x;
        }

        private void run_peirce(Excel.Range range)
        {
            //Get number of cells in range
            int N = range.Cells.Count;
            //Calculate mean
            double sum = 0.0;
            foreach (Excel.Range cell in range)
            {
                sum += cell.Value;
            }
            double mean = sum / N;

            //Calculate sample standard deviation
            double distance_sum_sq = 0;
            foreach (Excel.Range cell in range)
            {
                distance_sum_sq += Math.Pow(mean - cell.Value, 2);
            }
            double variance = distance_sum_sq / N;
            double std_dev = Math.Sqrt(variance);

            //Assume case of one doubtful observation to start
            int k = 1;
            //We will have one measured quantity
            int m = 1;
            int count_rejected = 0; 
            List<Excel.Range> outliers = new List<Excel.Range>();
            do
            {
                count_rejected = 0;
                //Obtain R corresponding to the number of measurements
                double max_z_score = get_peirce_cutoff(N, m, k);
                //If the Peirce cutoff is tiny, we are done
                if (max_z_score == 0)
                {
                    break;
                }
                //Calculate maximum allowable difference from the mean
                double max_difference_from_mean = max_z_score * std_dev;
                
                //Obtain |xi - mean| and look for outliers
                foreach (Excel.Range cell in range)
                {
                    bool already_outlier = false;
                    foreach (Excel.Range outlier in outliers)
                    {
                        if (outlier.Address.Equals(cell.Address))
                        {
                            already_outlier = true;
                        }
                    }
                    if (already_outlier)
                    {
                        continue;
                    }
                    else 
                    {
                        if (Math.Abs(cell.Value - mean) > max_difference_from_mean)
                        {
                            cell.Interior.Color = System.Drawing.Color.Red;
                            outliers.Add(cell);
                            count_rejected++;
                        }
                    }
                }
                k = k + count_rejected;
            } while (count_rejected > 0);
        }

        private List<double> run_peirce(double[] input_array)
        {
            //Get number of cells in range
            int N = input_array.Length;
            //Calculate mean
            double sum = 0.0;
            foreach (double d in input_array)
            {
                sum += d;
            }
            double mean = sum / N;

            //Calculate sample standard deviation
            double distance_sum_sq = 0;
            foreach (double d in input_array)
            {
                distance_sum_sq += Math.Pow(mean - d, 2);
            }
            double variance = distance_sum_sq / N;
            double std_dev = Math.Sqrt(variance);

            //Assume case of one doubtful observation to start
            int k = 1;
            //We will have one measured quantity
            int m = 1;
            int count_rejected = 0;
            List<double> outliers = new List<double>();
            do
            {
                count_rejected = 0;
                //Obtain R corresponding to the number of measurements
                double max_z_score = get_peirce_cutoff(N, m, k);
                //If the Peirce cutoff is tiny, we are done
                if (max_z_score == 0)
                {
                    break;
                }
                //Calculate maximum allowable difference from the mean
                double max_difference_from_mean = max_z_score * std_dev;

                //Obtain |xi - mean| and look for outliers
                foreach (double d in input_array)
                {
                    bool already_outlier = false;
                    foreach (double outlier in outliers)
                    {
                        if (outlier == d)
                        {
                            already_outlier = true;
                        }
                    }
                    if (already_outlier)
                    {
                        continue;
                    }
                    else
                    {
                        if (Math.Abs(d - mean) > max_difference_from_mean)
                        {
                            outliers.Add(d);
                            count_rejected++;
                        }
                    }
                }
                k = k + count_rejected;
            } while (count_rejected > 0);
            return outliers;
        }
    }
}
