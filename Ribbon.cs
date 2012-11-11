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

namespace DataDebug
{
    public partial class Ribbon
    {
        private bool toolHasNotRun = true; //this is to keep track of whether the tool has already run without having cleared the colorings
        List<TreeNode> nodes;        //This is a list holding all the TreeNodes in the Excel file
        TreeNode[][][] nodes_grid;   //This is a multi-dimensional array of TreeNodes that will hold all the TreeNodes -- stores the dependence graph
        double[][][][] impacts_grid; //This is a multi-dimensional array of doubles that will hold each cell's impact on each of the outputs
        double[][] min_max_delta_outputs; //This keeps the min and max delta for each output
        List<TreeNode> ranges;  //This is a list of all the ranges we have identified
        List<TreeNode> charts;  //This is a list of all the charts in the workbook
        private Regex[] regex_array;
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        /*
         * This method tries to break the worksheet apart into the composing ranges of cells. For now it only colors the ranges it identifies so that they can be seen visually.
         * If we decide to move forward with this approach, the identified ranges will be stored and analyzed afterward.
         * It looks at all numeric cells one by one. It looks at the cells to the right and underneath to see if they are of the same type.
         * It also relies on input from the user to say whether to favor a column-major or row-major format -- that is, whether to favor splitting up blocks of cells into columns or rows. 
         * Also, the user can specify if this method should analyze only at the selected range, or the entire worksheet. 
         */
        private void IdentifyRanges()
        {
            // Selects cells containing numeric constants
            Excel.Range specialCellConstants; // Will hold a range of all the numeric constants
            Excel.Worksheet activeWorksheet = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet;  //Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets[1]);

            //The specialcells method cannot be used on protected sheets, so if it is protected, we display a message
            if (activeWorksheet.ProtectContents)
            {
                MessageBox.Show("You must unprotect this worksheet to use this tool:\nReview tab -> Changes group -> Unprotect Sheet button\nYou may be prompted for a password.");
            }
            try  // Try is necessary in case there are no such cells to be selected (an exception is thrown if there are no such cells)
            {
                if (checkBox1.Checked) //Checks if the option for only analyzing the selection is checked
                {
                    specialCellConstants = Globals.ThisAddIn.Application.Selection as Excel.Range;
                }
                else //Otherwise it analyzes the entire range
                {
                    specialCellConstants = activeWorksheet.UsedRange.SpecialCells(Excel.XlCellType.xlCellTypeConstants, Microsoft.Office.Interop.Excel.XlSpecialCellsValue.xlNumbers);
                }

                int[] colColors = { 70756, 50756, 91156, 61156 };
                int[] rowColors = { 60756, 40756, 81156, 51156 };
                int colorBit1 = 0;
                //int colorBit2 = 0;
                int color1 = 0;
                //int color2 = 0;
                bool column_major = true;  // if vertical = 1, we assume columns are dominant. If vertical = 0, we assume rows are dominant. 
                if (dropDown1.SelectedItem.Label == "Column")
                {
                    column_major = true;
                }
                else
                {
                    column_major = false;
                }
                // Try to break apart special cell constants:
                // Foreach is a bit odd, but it starts from the top left of internal ranges. 
                // It moves ALONG ROWS until it gets to the end of the block; then moves to the beginning of the next row. 
                foreach (Excel.Range cell in specialCellConstants)
                {
                    bool cont = false; //Used to continue on to the following cell if the current one has already been checked
                    for (int i = 0; i < 4; i++)
                    {
                        if (column_major == true)
                        {
                            if (cell.Interior.Color == colColors[i])
                            {
                                cont = true;
                            }
                        }
                        else
                        {
                            if (cell.Interior.Color == rowColors[i])
                            {
                                cont = true;
                            }
                        }
                    }
                    if (cont == true)
                    {
                        continue;
                    }

                    //    if (cell.Interior.Color == 70756 || cell.Interior.Color == 50756 || cell.Interior.Color == 91156 || cell.Interior.Color == 61156)
                    //  {
                    //    continue;
                    //}
                    //cell.Interior.Color = System.Drawing.Color.BlanchedAlmond;
                    //  MessageBox.Show("Colored cell " + cell.Address);
                    Excel.Range cellRight = cell.get_Offset(0, 1);
                    Excel.Range cellUnder = cell.get_Offset(1, 0);
                    if (column_major == true)
                    {
                        if (typesMatch(cell, cellUnder))
                        {
                            while (typesMatch(cell, cellUnder) &&
                                Globals.ThisAddIn.Application.Intersect(cellUnder, specialCellConstants, Type.Missing,
                                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing) != null)
                            {
                                if (colorBit1 == 0)
                                    color1 = 70756;
                                else
                                    color1 = 50756;
                                cell.Interior.Color = color1;
                                cellUnder.Interior.Color = color1;// System.Drawing.Color.Blue;
                                cellUnder = cellUnder.get_Offset(1, 0);
                            }
                            colorBit1 = (colorBit1 + 1) % 2;
                        }
                        /**
                    else
                    {
                        if (typesMatch(cell, cellRight))
                        {
                            while (typesMatch(cell, cellRight) && 
                            Globals.ThisAddIn.Application.Intersect(cellRight,specialCellConstants, Type.Missing, 
                            Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,Type.Missing, 
                            Type.Missing,Type.Missing,Type.Missing, Type.Missing,Type.Missing, Type.Missing, 
                            Type.Missing, Type.Missing,Type.Missing, Type.Missing,Type.Missing, Type.Missing,
                            Type.Missing, Type.Missing,Type.Missing, Type.Missing,Type.Missing, Type.Missing, Type.Missing) != null)
                            {
                                if (colorBit2 == 0)
                                    color2 = 91156;
                                else
                                    color2 = 61156;
                                cell.Interior.Color = color2;
                                cellRight.Interior.Color = color2; // System.Drawing.Color.HotPink;
                                cellRight = cellRight.get_Offset(0, 1);
                            }
                            colorBit2 = (colorBit2 + 1) % 2;
                        }
                        else
                        {
                            //cc = MessageBox.Show("CALLED METHOD: Right cell DOES NOT match type.");
                        }

                    }
                         * **/
                    }
                    else
                    {
                        if (typesMatch(cell, cellRight))
                        {
                            while (typesMatch(cell, cellRight) &&
                                Globals.ThisAddIn.Application.Intersect(cellRight, specialCellConstants, Type.Missing,
                                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing) != null)
                            {
                                if (colorBit1 == 0)
                                    color1 = 60756;
                                else
                                    color1 = 40756;
                                cell.Interior.Color = color1;
                                cellRight.Interior.Color = color1; // System.Drawing.Color.Blue;
                                cellRight = cellRight.get_Offset(0, 1);
                            }
                            colorBit1 = (colorBit1 + 1) % 2;
                        }
                        /**
                        else
                        {
                            if (typesMatch(cell, cellUnder))
                            {
                                while (typesMatch(cell, cellUnder) &&
                                Globals.ThisAddIn.Application.Intersect(cellUnder, specialCellConstants, Type.Missing,
                                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing) != null)
                                {
                                    if (colorBit2 == 0)
                                        color2 = 81156;
                                    else
                                        color2 = 51156;
                                    cell.Interior.Color = color2;
                                    cellUnder.Interior.Color = color2; // System.Drawing.Color.HotPink;
                                    cellUnder = cellUnder.get_Offset(1, 0);
                                }
                                colorBit2 = (colorBit2 + 1) % 2;
                            }
                            else
                            {
                                //cc = MessageBox.Show("CALLED METHOD: Right cell DOES NOT match type.");
                            }

                        } **/
                    }

                    //i++;
                    //if (i>20)
                    //    break;
                }
            }
            catch
            {
            }

            //This will color all cells containing formulas crimson
            try  //Necessary in case there are no matching cells
            {
                Excel.Range specialCellFormulas = activeWorksheet.UsedRange.SpecialCells(Excel.XlCellType.xlCellTypeFormulas, Type.Missing); //.xlCellTypeAllFormatConditions);
                //specialCellFormulas.Interior.Color = System.Drawing.Color.Crimson;
            }
            catch { }
        }

        /*
         * Checks if the types of two cells match
         * If the types match, it returns true; otherwise it returns false
         * This method is used for trying to break up a worksheet into separate ranges of cells
         */
        private Boolean typesMatch(Excel.Range cell1, Excel.Range cell2)
        {
            if (cell2.get_Value() != null)
            {
                if (cell2.get_Value().GetType() == cell1.get_Value().GetType())
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
            else
            {
                return false;
            }
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
         * This is a recursive method for propagating the weights down the nodes in the tree
         * All outputs have weight 1. Their n children have weight 1/n, and so forth. 
         */
        private void propagateWeightUp(TreeNode node, double weight_passed_up)
        {
            if (!node.hasChildren())
            {
                return;
            }
            else
            {
                foreach (TreeNode child in node.getChildren())
                {
                    child.setWeight(child.getWeight() + weight_passed_up);
                    propagateWeightUp(child, 1.0);
                }
            }
        }

        /*
         * This method constructs the dependency graph from the worksheet.
         * It analyzes formulas and looks for references to cells or ranges of cells.
         * It also looks for any charts, and adds those to the dependency graph as well. 
         * After the dependency graph is constructed, we use it to determine and propagate weights to all nodes in the graph. 
         * In the end, a text representation of the dependency graph is given in GraphViz format. It includes the entire graph and the weights of the nodes.
         */
        private void constructTree()
        {
            //Excel.Worksheet activeWorksheet = ((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet);
            //List<TreeNode> nodes = new List<TreeNode>();        //This is a list holding all the TreeNodes in the Excel file
            Excel.Range analysisRange = null; //This keeps track of the range to be analyzed - it is either the user's selection or the whole workbook
            Excel.Range[] analysisRanges = new Excel.Range[Globals.ThisAddIn.Application.Worksheets.Count]; //This keeps track of the range to be analyzed in every worksheet of the workbook
            if (checkBox1.Checked) //if "Use selection" box is checked
            {
                analysisRange = Globals.ThisAddIn.Application.Selection as Excel.Range;
                Excel.Range number_cells = null;
                Excel.Range formula_cells = null;
                try
                {
                    number_cells = analysisRange.SpecialCells(Excel.XlCellType.xlCellTypeConstants, Microsoft.Office.Interop.Excel.XlSpecialCellsValue.xlNumbers);
                }
                catch 
                {
                    try
                    {
                        number_cells = analysisRange.SpecialCells(Excel.XlCellType.xlCellTypeFormulas, Type.Missing);
                    }
                    catch {}
                }
                try
                {
                    formula_cells = analysisRange.SpecialCells(Excel.XlCellType.xlCellTypeFormulas, Type.Missing);
                }
                catch
                {
                    try
                    {
                        formula_cells = analysisRange.SpecialCells(Excel.XlCellType.xlCellTypeConstants, Microsoft.Office.Interop.Excel.XlSpecialCellsValue.xlNumbers);
                    }
                    catch {}
                }
                try
                {
                    analysisRange = Globals.ThisAddIn.Application.Union(number_cells, formula_cells, Type.Missing,
                                            Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                            Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                            Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                            Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                }
                catch
                {
                    MessageBox.Show("No cells to analyze.");
                }
            }
            else  //if "Use selection" box is not checked
            {
                analysisRange = null; // activeWorksheet.UsedRange; // Globals.ThisAddIn.Application.Selection as Excel.Range;
                int worksheet_index = 0; // keeps track of which worksheet we are currently examining
                foreach (Excel.Worksheet ws in Globals.ThisAddIn.Application.Worksheets)
                {
                    Excel.Range number_cells = null;
                    Excel.Range formula_cells = null;
                    try
                    {
                        number_cells = ws.UsedRange.SpecialCells(Excel.XlCellType.xlCellTypeConstants, Microsoft.Office.Interop.Excel.XlSpecialCellsValue.xlNumbers);
                    }
                    catch
                    {
                        try
                        {
                            number_cells = ws.UsedRange.SpecialCells(Excel.XlCellType.xlCellTypeFormulas, Type.Missing);
                        }
                        catch { }
                    }
                    try
                    {
                        formula_cells = ws.UsedRange.SpecialCells(Excel.XlCellType.xlCellTypeFormulas, Type.Missing);
                    }
                    catch
                    {
                        try
                        {
                            formula_cells = ws.UsedRange.SpecialCells(Excel.XlCellType.xlCellTypeConstants, Microsoft.Office.Interop.Excel.XlSpecialCellsValue.xlNumbers);
                        }
                        catch { }
                    }
                    try
                    {
                        analysisRanges[worksheet_index] = Globals.ThisAddIn.Application.Union(
                                            number_cells, //activeWorksheet.UsedRange.SpecialCells(Excel.XlCellType.xlCellTypeConstants, Microsoft.Office.Interop.Excel.XlSpecialCellsValue.xlNumbers), //activeWorksheet.UsedRange;
                                            formula_cells, //activeWorksheet.UsedRange.SpecialCells(Excel.XlCellType.xlCellTypeFormulas, Type.Missing), 
                                            Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                            Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                            Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                            Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                    }
                    catch
                    {
                        analysisRanges[worksheet_index] = null;
                        //MessageBox.Show("No cells to analyze in worksheet " + ws.Name +".");
                    }
                    worksheet_index++;
                }
            }
            System.Diagnostics.Stopwatch swTree = new System.Diagnostics.Stopwatch();
            swTree.Start();
            //First we create nodes for every non-null cell; then we will operate on these node objects, connecting them in the tree, etc. 
            //This includes cells that contain constants and formulas
            if (analysisRange != null) //if we are only analyzing the user's selection, create nodes only for the selection
            {
                foreach (Excel.Range cell in analysisRange)
                {
                    //TODO Test the functionality of selecting only a part of the worksheet to analyze. 
                    //MessageBox.Show(cell.Worksheet.Name + ": " + cell.Address);
                    if (cell.Value != null)
                    {
                        TreeNode n = new TreeNode(cell.Address, cell.Worksheet.Name);  //Create a TreeNode for every cell with the name being the cell's address and set the node's worksheet appropriately
                        if (toggle_array_storage.Checked)
                        {
                            nodes_grid[cell.Worksheet.Index - 1][cell.Row - 1][cell.Column - 1] = n;
                        }
                        else  //Toggle_array_storage unchecked
                        {
                            nodes.Add(n);
                        }
                    }
                }
                return;
            }
            else  //if we are analyzing the entire workbook, create nodes for the non-null cells in all the worksheets
            {
                // Go through every worksheet
                foreach (Excel.Range worksheet_range in analysisRanges)
                {
                    // Go through every cell of every worksheet
                    if (worksheet_range != null)
                    {
                        foreach (Excel.Range cell in worksheet_range)
                        {
                            if (cell.Value != null)
                            {
                                TreeNode n = new TreeNode(cell.Address, cell.Worksheet.Name);  //Create a TreeNode for every cell with the name being the cell's address and set the node's worksheet appropriately
                                if (toggle_array_storage.Checked)
                                {
                                    try
                                    {
                                        nodes_grid[cell.Worksheet.Index - 1][cell.Row - 1][cell.Column - 1] = n;
                                    }
                                    catch
                                    {
                                        cell.Interior.Color = System.Drawing.Color.Purple;
                                    }
                                }
                                else  //Toggle_array_storage unchecked
                                {
                                    nodes.Add(n);
                                }
                            }
                        }
                    }
                    else
                    {
                        continue;
                    }
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
                //MessageBox.Show(worksheet.Name);
            }

            if (analysisRange != null) //if we are only analyzing the user's selection
            {
                //TODO Future work - analyze a user's selected range
                //foreach (Excel.Range c in analysisRange)
                //{
                //    if (c.HasFormula)
                //    {
                //        TreeNode formula_cell = null;
                //        //Look for the node object for the current cell in the list of TreeNodes
                //        foreach (TreeNode n in nodes)
                //        {
                //            if (n.getName() == c.Address && n.getWorksheet() == c.Worksheet.Name)
                //            {
                //                formula_cell = n;
                //            }
                //            else
                //            {
                //                continue;
                //            }
                //        }

                //        string formula = c.Formula;  //The formula contained in the cell
                //        MatchCollection matchedRanges = null;
                //        MatchCollection matchedCells = null;
                //        int ws_index = 1;
                //        foreach (string s in worksheet_names)
                //        {
                //            string worksheet_name = s.Replace("+", @"\+").Replace("^", @"\^").Replace("$", @"\$").Replace(".", @"\."); //Escape certain characters in the regular expression
                //            //First look for range references of the form 'worksheet_name'!A1:A10 in the formula (with quotation marks around the name)
                //            if (toggle_compile_regex.Checked)
                //            {
                //                Regex regex = new Regex(@"('" + worksheet_name + @"'!\$?[A-Z]+\$?[1-9]\d*:\$?[A-Z]+\$?[1-9]\d*)", RegexOptions.Compiled);
                //                matchedRanges = regex.Matches(formula);
                //            }
                //            else
                //            {
                //                matchedRanges = Regex.Matches(formula, @"('" + worksheet_name + @"'!\$?[A-Z]+\$?[1-9]\d*:\$?[A-Z]+\$?[1-9]\d*)");  //A collection of all the references in the formula to ranges in the particular worksheet; each item is a range reference of the form 'worksheet_name'!A1:A10
                //            }
                //            foreach (Match match in matchedRanges)
                //            {
                //                formula = formula.Replace(match.Value, "");
                //                string ws_name = worksheet_name; //match.Value.Substring(1, match.Value.LastIndexOf("!") - 2); // Get the name of the worksheet being referenced
                //                string range_coordinates = match.Value.Substring(match.Value.LastIndexOf("!") + 1); //match.Value.Replace("'" + ws_name + "'!", "");
                //                string[] endCells = range_coordinates.Split(':');     //Split up each matched range into the start and end cells of the range
                //                TreeNode range = null;
                //                //Try to find the range in existing TreeNodes
                //                if (toggle_array_storage.Checked)
                //                {
                //                    foreach (TreeNode n in ranges)
                //                    {
                //                        if (n.getName().Replace("$", "") == endCells[0].Replace("$", "") + "_to_" + endCells[1].Replace("$", "") && n.getWorksheet() == ws_name)
                //                        {
                //                            //MessageBox.Show("Found node:" + ws_name + "_" + endCells[0] + ":" + endCells[1]);
                //                            range = n;
                //                        }
                //                        else
                //                        {
                //                            continue;
                //                        }
                //                    }
                //                }
                //                else
                //                {
                //                    foreach (TreeNode n in nodes)
                //                    {
                //                        if (n.getName().Replace("$", "") == endCells[0].Replace("$", "") + "_to_" + endCells[1].Replace("$", "") && n.getWorksheet() == ws_name)
                //                        {
                //                            //MessageBox.Show("Found node:" + ws_name + "_" + endCells[0] + ":" + endCells[1]);
                //                            range = n;
                //                        }
                //                        else
                //                        {
                //                            continue;
                //                        }
                //                    }
                //                }
                //                //If it does not exist, create it
                //                if (range == null)
                //                {
                //                    range = new TreeNode(endCells[0].Replace("$", "") + "_to_" + endCells[1].Replace("$", ""), ws_name);
                //                    //MessageBox.Show("Created range node:" + ws_name + "_" + endCells[0] + ":" + endCells[1]);
                //                    if (toggle_array_storage.Checked)
                //                    {
                //                        ranges.Add(range);
                //                    }
                //                    else  //Toggle_array_storage unchecked
                //                    {
                //                        nodes.Add(range);
                //                    }
                //                }
                //                formula_cell.addParent(range);
                //                range.addChild(formula_cell);
                //                //Add each cell contained in the range to the dependencies
                //                foreach (Excel.Range cellInRange in Globals.ThisAddIn.Application.Worksheets[ws_index].Range[endCells[0], endCells[1]])
                //                {
                //                    TreeNode input_cell = null;
                //                    //Find the node object for the current cell in the list of TreeNodes
                //                    foreach (TreeNode node in nodes)
                //                    {
                //                        if (node.getName().Replace("$", "") == cellInRange.Address.Replace("$", "") && node.getWorksheet() == cellInRange.Worksheet.Name)
                //                        {
                //                            input_cell = node;
                //                        }
                //                        else
                //                            continue;
                //                    }
                //                    //If it wasn't found, then it is blank, and we have to create a TreeNode for it
                //                    if (input_cell == null)
                //                    {
                //                        input_cell = new TreeNode(cellInRange.Address, cellInRange.Worksheet.Name);
                //                        nodes.Add(input_cell);
                //                    }

                //                    //Update the dependencies
                //                    range.addParent(input_cell);
                //                    input_cell.addChild(range);
                //                }
                //            }

                //            //Next look for range references of the form worksheet_name!A1:A10 in the formula (no quotation marks around the name)
                //            matchedRanges = Regex.Matches(formula, @"(" + worksheet_name + @"!\$?[A-Z]+\$?[1-9]\d*:\$?[A-Z]+\$?[1-9]\d*)");  //A collection of all the range references in the formula; each item is a range reference such as A1:A10
                //            foreach (Match match in matchedRanges)
                //            {
                //                formula = formula.Replace(match.Value, "");
                //                string ws_name = worksheet_name; // match.Value.Substring(0, match.Value.LastIndexOf("!")); // Get the name of the worksheet being referenced
                //                string range_coordinates = match.Value.Substring(match.Value.LastIndexOf("!") + 1);  //match.Value.Replace(ws_name + "!", "");
                //                string[] endCells = range_coordinates.Split(':');     //Split up each matched range into the start and end cells of the range
                //                TreeNode range = null;
                //                //Try to find the range in existing TreeNodes
                //                foreach (TreeNode n in nodes)
                //                {
                //                    if (n.getName().Replace("$", "") == endCells[0].Replace("$", "") + "_to_" + endCells[1].Replace("$", "") && n.getWorksheet() == ws_name)
                //                    {
                //                        //MessageBox.Show("Found node:" + ws_name + "_" + endCells[0] + ":" + endCells[1]);
                //                        range = n;
                //                    }
                //                    else
                //                    {
                //                        continue;
                //                    }
                //                }
                //                //If it does not exist, create it
                //                if (range == null)
                //                {
                //                    range = new TreeNode(endCells[0].Replace("$", "") + "_to_" + endCells[1].Replace("$", ""), ws_name);
                //                    //MessageBox.Show("Created node:" + ws_name + "_" + endCells[0] + ":" + endCells[1]);
                //                    nodes.Add(range);
                //                }
                //                formula_cell.addParent(range);
                //                range.addChild(formula_cell);
                //                //Add each cell contained in the range to the dependencies
                //                foreach (Excel.Range cellInRange in Globals.ThisAddIn.Application.Worksheets[ws_index].Range[endCells[0], endCells[1]])
                //                {
                //                    TreeNode input_cell = null;
                //                    //Find the node object for the current cell in the list of TreeNodes
                //                    foreach (TreeNode node in nodes)
                //                    {
                //                        if (node.getName().Replace("$", "") == cellInRange.Address.Replace("$", "") && node.getWorksheet() == cellInRange.Worksheet.Name)
                //                        {
                //                            input_cell = node;
                //                        }
                //                        else
                //                            continue;
                //                    }
                //                    //If it wasn't found, then it is blank, and we have to create a TreeNode for it
                //                    if (input_cell == null)
                //                    {
                //                        input_cell = new TreeNode(cellInRange.Address, cellInRange.Worksheet.Name);
                //                        nodes.Add(input_cell);
                //                    }

                //                    //Update the dependencies
                //                    range.addParent(input_cell);
                //                    input_cell.addChild(range);
                //                }
                //            }

                //            // Now we look for references of the kind 'worksheet_name'!A1 (with quotation marks)
                //            matchedCells = Regex.Matches(formula, @"('" + worksheet_name + @"'!\$?[A-Z]+\$?[1-9]\d*)"); //matchedCells is a collection of all the references in the formula to cells in the specific worksheet, where the reference has the form 'worksheet_name'!A1
                //            foreach (Match match in matchedCells)
                //            {
                //                formula = formula.Replace(match.Value, "");
                //                string ws_name = worksheet_name; // match.Value.Substring(1, match.Value.LastIndexOf("!") - 2); // Get the name of the worksheet being referenced
                //                string cell_coordinates = match.Value.Substring(match.Value.LastIndexOf("!") + 1);

                //                TreeNode input_cell = null;
                //                //Find the node object for the current cell in the list of TreeNodes
                //                foreach (TreeNode node in nodes)
                //                {
                //                    if (node.getName().Replace("$", "") == cell_coordinates.Replace("$", "") && node.getWorksheet() == ws_name)
                //                    {
                //                        input_cell = node;
                //                    }
                //                    else
                //                    {
                //                        continue;
                //                    }
                //                }
                //                //If it wasn't found, then it is blank, and we have to create a TreeNode for it
                //                if (input_cell == null)
                //                {
                //                    input_cell = new TreeNode(cell_coordinates.Replace("$", ""), ws_name);
                //                    nodes.Add(input_cell);
                //                }

                //                //Update the dependencies
                //                formula_cell.addParent(input_cell);
                //                input_cell.addChild(formula_cell);
                //            }

                //            //Lastly we look for references of the kind worksheet_name!A1 (without quotation marks)
                //            matchedCells = Regex.Matches(formula, @"(" + worksheet_name + @"!\$?[A-Z]+\$?[1-9]\d*)");
                //            foreach (Match match in matchedCells)
                //            {
                //                formula = formula.Replace(match.Value, "");
                //                string ws_name = worksheet_name; //match.Value.Substring(0, match.Value.LastIndexOf("!")); // Get the name of the worksheet being referenced
                //                string cell_coordinates = match.Value.Substring(match.Value.LastIndexOf("!") + 1);

                //                TreeNode input_cell = null;
                //                //Find the node object for the current cell in the list of TreeNodes
                //                foreach (TreeNode node in nodes)
                //                {
                //                    if (node.getName().Replace("$", "") == cell_coordinates.Replace("$", "") && node.getWorksheet() == ws_name)
                //                    {
                //                        input_cell = node;
                //                    }
                //                    else
                //                    {
                //                        continue;
                //                    }
                //                }
                //                //If it wasn't found, then it is blank, and we have to create a TreeNode for it
                //                if (input_cell == null)
                //                {
                //                    input_cell = new TreeNode(cell_coordinates.Replace("$", ""), ws_name);
                //                    nodes.Add(input_cell);
                //                }

                //                //Update the dependencies
                //                formula_cell.addParent(input_cell);
                //                input_cell.addChild(formula_cell);
                //            }
                //            ws_index++;
                //        }
                //        // Now we look for range references and cell references not involving worksheet references
                //        string patternRange = @"(\$?[A-Z]+\$?[1-9]\d*:\$?[A-Z]+\$?[1-9]\d*)";  //Regex for matching range references in formulas such as A1:A10, or $A$1:$A$10 etc.
                //        string patternCell = @"(\$?[A-Z]+\$?[1-9]\d*)";        //Regex for matching single cell references such as A1 or $A$1, etc. 

                //        //First look for range references in the formula
                //        matchedRanges = Regex.Matches(formula, patternRange);  //A collection of all the range references in the formula; each item is a range reference such as A1:A10
                //        List<Excel.Range> rangeList = new List<Excel.Range>();
                //        foreach (Match match in matchedRanges)
                //        {
                //            formula = formula.Replace(match.Value, "");
                //            string[] endCells = match.Value.Split(':');     //Split up each matched range into the start and end cells of the range
                //            TreeNode range = null;
                //            //Try to find the range in existing TreeNodes
                //            foreach (TreeNode n in nodes)
                //            {
                //                if (n.getName().Replace("$", "") == endCells[0].Replace("$", "") + "_to_" + endCells[1].Replace("$", "") && n.getWorksheet() == c.Worksheet.Name)
                //                {
                //                    range = n;
                //                }
                //                else
                //                {
                //                    continue;
                //                }
                //            }
                //            //If it does not exist, create it
                //            if (range == null)
                //            {
                //                //MessageBox.Show("Created range node:" + c.Worksheet.Name + "_" + endCells[0] + ":" + endCells[1]);
                //                range = new TreeNode(endCells[0].Replace("$", "") + "_to_" + endCells[1].Replace("$", ""), c.Worksheet.Name);
                //                nodes.Add(range);
                //            }
                //            formula_cell.addParent(range);
                //            range.addChild(formula_cell);
                //            //Add each cell contained in the range to the dependencies
                //            foreach (Excel.Range cellInRange in c.Worksheet.Range[endCells[0], endCells[1]])
                //            {
                //                TreeNode input_cell = null;
                //                //Find the node object for the current cell in the list of TreeNodes
                //                foreach (TreeNode node in nodes)
                //                {
                //                    if (node.getName().Replace("$", "") == cellInRange.Address.Replace("$", "") && node.getWorksheet() == c.Worksheet.Name)
                //                    {
                //                        input_cell = node;
                //                    }
                //                    else
                //                        continue;
                //                }
                //                //If it wasn't found, then it is blank, and we have to create a TreeNode for it
                //                if (input_cell == null)
                //                {
                //                    input_cell = new TreeNode(cellInRange.Address, cellInRange.Worksheet.Name);
                //                    nodes.Add(input_cell);
                //                }

                //                //Update the dependencies
                //                range.addParent(input_cell);
                //                input_cell.addChild(range);
                //            }
                //        }

                //        matchedCells = Regex.Matches(formula, patternCell);  //matchedCells is a collection of all the cells that are referenced by the formula
                //        foreach (Match m in matchedCells)
                //        {
                //            TreeNode input_cell = null;
                //            //MessageBox.Show(m.Value);
                //            //Find the node object for the current cell in the list of TreeNodes
                //            foreach (TreeNode node in nodes)
                //            {
                //                if (node.getName().Replace("$", "") == m.Value.Replace("$", "") && node.getWorksheet() == c.Worksheet.Name)
                //                {
                //                    input_cell = node;
                //                }
                //                else
                //                {
                //                    continue;
                //                }
                //            }
                //            //If it wasn't found, then it is blank, and we have to create a TreeNode for it
                //            if (input_cell == null)
                //            {
                //                input_cell = new TreeNode(m.Value.Replace("$", ""), c.Worksheet.Name);
                //                nodes.Add(input_cell);
                //            }

                //            //Update the dependencies
                //            formula_cell.addParent(input_cell);
                //            input_cell.addChild(formula_cell);
                //        }
                //    }
                //}
                ///**
                //foreach (Excel.Range c in analysisRange)
                //{
                //    if (c.HasFormula)
                //    {
                //        TreeNode formula_cell = null;
                //        //Look for the node object for the current cell in the list of TreeNodes
                //        foreach (TreeNode n in nodes)
                //        {
                //            if (n.getName() == c.Address && n.getWorksheet() == c.Worksheet.Name)
                //            {
                //                formula_cell = n;
                //            }
                //            else
                //            {
                //                continue;
                //            }
                //        }

                //        string patternRange = @"(\$?[A-Z]+\$?[1-9]\d*:\$?[A-Z]+\$?[1-9]\d*)";  //Regex for matching range references in formulas such as A1:A10, or $A$1:$A$10 etc.
                //        string patternCell = @"(\$?[A-Z]+\$?[1-9]\d*)";        //Regex for matching single cell references such as A1 or $A$1, etc. 
                //        string formula = c.Formula;  //The formula contained in the cell

                //        //First look for range references in the formula
                //        MatchCollection matchedRanges = Regex.Matches(formula, patternRange);  //A collection of all the range references in the formula; each item is a range reference such as A1:A10
                //        List<Excel.Range> rangeList = new List<Excel.Range>();
                //        foreach (Match match in matchedRanges)
                //        {
                //            formula = formula.Replace(match.Value, "");
                //            string[] endCells = match.Value.Split(':');     //Split up each matched range into the start and end cells of the range
                //            TreeNode range = null;
                //            //Try to find the range in existing TreeNodes
                //            foreach (TreeNode n in nodes)
                //            {
                //                if (n.getName() == endCells[0].Replace("$", "") + "_to_" + endCells[1].Replace("$", ""))
                //                {
                //                    range = n;
                //                }
                //                else
                //                {
                //                    continue;
                //                }
                //            }
                //            //If it does not exist, create it
                //            if (range == null)
                //            {
                //                //TODO FIX WORKSHEET ARGUMENT:
                //                range = new TreeNode(endCells[0].Replace("$", "") + "_to_" + endCells[1].Replace("$", ""), c.Worksheet.Name);
                //                nodes.Add(range);
                //            }
                //            formula_cell.addParent(range);
                //            range.addChild(formula_cell);
                //            //Add each cell contained in the range to the dependencies
                //            foreach (Excel.Range cellInRange in activeWorksheet.Range[endCells[0], endCells[1]])
                //            {
                //                TreeNode input_cell = null;
                //                //Find the node object for the current cell in the list of TreeNodes
                //                foreach (TreeNode node in nodes)
                //                {
                //                    if (node.getName() == cellInRange.Address)
                //                    {
                //                        input_cell = node;
                //                    }
                //                    else
                //                        continue;
                //                }
                //                //If it wasn't found, then it is blank, and we have to create a TreeNode for it
                //                if (input_cell == null)
                //                {
                //                    input_cell = new TreeNode(cellInRange.Address, cellInRange.Worksheet.Name);
                //                    nodes.Add(input_cell);
                //                }

                //                //Update the dependencies
                //                range.addParent(input_cell);
                //                input_cell.addChild(range);
                //            }
                //        }

                //        MatchCollection matchedCells = Regex.Matches(formula, patternCell);  //matchedCells is a collection of all the cells that are referenced by the formula
                //        foreach (Match m in matchedCells)
                //        {
                //            TreeNode input_cell = null;
                //            MessageBox.Show(m.Value);
                //            //Find the node object for the current cell in the list of TreeNodes
                //            foreach (TreeNode node in nodes)
                //            {
                //                if (node.getName().Replace("$", "") == m.Value.Replace("$", ""))
                //                {
                //                    input_cell = node;
                //                }
                //                else
                //                {
                //                    continue;
                //                }
                //            }

                //            //Update the dependencies
                //            formula_cell.addParent(input_cell);
                //            input_cell.addChild(formula_cell);
                //        }
                //    }
                //}
                //**/
            }
            else  // if we are analyzing the entire workbook
            {
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
                                if (toggle_array_storage.Checked)
                                {
                                    //Check if this cell's coordinates are within the bounds of the used range, otherwise there will be an index out of bounds error
                                    if (c.Column <= c.Worksheet.UsedRange.Columns.Count && c.Row <= c.Worksheet.UsedRange.Rows.Count)
                                    {
                                        //if a TreeNode exists for this cell already
                                        if (nodes_grid[c.Worksheet.Index - 1][c.Row - 1][c.Column - 1] != null)
                                        {
                                            formula_cell = nodes_grid[c.Worksheet.Index - 1][c.Row - 1][c.Column - 1];
                                        }
                                    }
                                }
                                else //toggle_array_storage not checked
                                {
                                    foreach (TreeNode n in nodes)
                                    {
                                        if (n.getName() == c.Address && n.getWorksheet() == c.Worksheet.Name)
                                        {
                                            formula_cell = n;
                                        }
                                        else
                                        {
                                            continue;
                                        }
                                    }
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
                                //find_references(formula_cell, formula);
                                
                                if (toggle_reporting.Checked)
                                {
                                    MessageBox.Show(formula);
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
                                    if (toggle_reporting.Checked)
                                    {
                                        MessageBox.Show("OK 1");
                                    }
                                    foreach (Match match in matchedRanges)
                                    {
                                        formula = formula.Replace(match.Value, "");
                                        string ws_name = worksheet_name; // match.Value.Substring(1, match.Value.LastIndexOf("!") - 2); // Get the name of the worksheet being referenced
                                        string range_coordinates = match.Value.Substring(match.Value.LastIndexOf("!") + 1); //match.Value.Replace("'" + ws_name + "'!", "");
                                        string[] endCells = range_coordinates.Split(':');     //Split up each matched range into the start and end cells of the range
                                        TreeNode range = null;
                                        //Try to find the range in existing TreeNodes
                                        if (toggle_array_storage.Checked)
                                        {
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
                                        }
                                        else //toggle_array_storage not checked
                                        {
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
                                        }
                                        //If it was not found, create it
                                        if (range == null)
                                        {
                                            range = new TreeNode(endCells[0].Replace("$", "") + "_to_" + endCells[1].Replace("$", ""), ws_name);
                                            //MessageBox.Show("Created range node:" + ws_name + "_" + endCells[0] + ":" + endCells[1]);
                                            if (toggle_array_storage.Checked)
                                            {
                                                ranges.Add(range);
                                            }
                                            else //toggle_array_storage not checked
                                            {
                                                nodes.Add(range);
                                            }
                                        }
                                        formula_cell.addParent(range);
                                        range.addChild(formula_cell);
                                        //Add each cell contained in the range to the dependencies
                                        foreach (Excel.Range cellInRange in Globals.ThisAddIn.Application.Worksheets[ws_index].Range[endCells[0], endCells[1]])
                                        {
                                            TreeNode input_cell = null;
                                            //Find the node object for the current cell in the existing TreeNodes
                                            if (toggle_array_storage.Checked)
                                            {
                                                //Check if this cell's coordinates are within the bounds of the used range, otherwise there will be an index out of bounds error
                                                if (cellInRange.Column <= cellInRange.Worksheet.UsedRange.Columns.Count && cellInRange.Row <= cellInRange.Worksheet.UsedRange.Rows.Count)
                                                {
                                                    //if a TreeNode exists for this cell already
                                                    if (nodes_grid[cellInRange.Worksheet.Index - 1][cellInRange.Row - 1][cellInRange.Column - 1] != null)
                                                    {
                                                        input_cell = nodes_grid[cellInRange.Worksheet.Index - 1][cellInRange.Row - 1][cellInRange.Column - 1];
                                                    }
                                                }
                                            }
                                            else  //toggle_array_storage not checked
                                            {
                                                foreach (TreeNode node in nodes)
                                                {
                                                    if (node.getName().Replace("$", "") == cellInRange.Address.Replace("$", "") && node.getWorksheet() == cellInRange.Worksheet.Name)
                                                    {
                                                        input_cell = node;
                                                    }
                                                    else
                                                        continue;
                                                }
                                            }
                                            //If it wasn't found, then it is blank, and we have to create a TreeNode for it
                                            if (input_cell == null)
                                            {
                                                input_cell = new TreeNode(cellInRange.Address, cellInRange.Worksheet.Name);
                                                if (toggle_array_storage.Checked)
                                                {
                                                    //Check if this cell's coordinates are within the bounds of the used range, otherwise there will be an index out of bounds error
                                                    if (cellInRange.Column <= cellInRange.Worksheet.UsedRange.Columns.Count && cellInRange.Row <= cellInRange.Worksheet.UsedRange.Rows.Count)
                                                    {
                                                        nodes_grid[cellInRange.Worksheet.Index - 1][cellInRange.Row - 1][cellInRange.Column - 1] = input_cell;
                                                    }
                                                }
                                                else  //toggle_array_storage not checked
                                                {
                                                    nodes.Add(input_cell);
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
                                    if (toggle_reporting.Checked)
                                    {
                                        MessageBox.Show("OK 2");
                                    }
                                    foreach (Match match in matchedRanges)
                                    {
                                        formula = formula.Replace(match.Value, "");
                                        string ws_name = worksheet_name; // match.Value.Substring(0, match.Value.LastIndexOf("!")); // Get the name of the worksheet being referenced
                                        string range_coordinates = match.Value.Substring(match.Value.LastIndexOf("!") + 1);  //match.Value.Replace(ws_name + "!", "");
                                        string[] endCells = range_coordinates.Split(':');     //Split up each matched range into the start and end cells of the range
                                        TreeNode range = null;
                                        //Try to find the range in existing TreeNodes
                                        if (toggle_array_storage.Checked)
                                        {
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
                                        }
                                        else //toggle_array_storage not checked
                                        {
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
                                        }
                                        //If it does not exist, create it
                                        if (range == null)
                                        {
                                            range = new TreeNode(endCells[0].Replace("$", "") + "_to_" + endCells[1].Replace("$", ""), ws_name);
                                            //MessageBox.Show("Created node:" + ws_name + "_" + endCells[0] + ":" + endCells[1]);
                                            if (toggle_array_storage.Checked)
                                            {
                                                ranges.Add(range);
                                            }
                                            else //toggle_array_storage not checked
                                            {
                                                nodes.Add(range);
                                            }
                                        }
                                        formula_cell.addParent(range);
                                        range.addChild(formula_cell);
                                        //Add each cell contained in the range to the dependencies
                                        foreach (Excel.Range cellInRange in Globals.ThisAddIn.Application.Worksheets[ws_index].Range[endCells[0], endCells[1]])
                                        {
                                            TreeNode input_cell = null;
                                            //Find the node object for the current cell in the existing TreeNodes
                                            if (toggle_array_storage.Checked)
                                            {
                                                //Check if this cell's coordinates are within the bounds of the used range, otherwise there will be an index out of bounds error
                                                if (cellInRange.Column <= cellInRange.Worksheet.UsedRange.Columns.Count && cellInRange.Row <= cellInRange.Worksheet.UsedRange.Rows.Count)
                                                {
                                                    //if a TreeNode exists for this cell already
                                                    if (nodes_grid[cellInRange.Worksheet.Index - 1][cellInRange.Row - 1][cellInRange.Column - 1] != null)
                                                    {
                                                        input_cell = nodes_grid[cellInRange.Worksheet.Index - 1][cellInRange.Row - 1][cellInRange.Column - 1];
                                                    }
                                                }
                                            }
                                            else  //toggle_array_storage not checked
                                            {
                                                foreach (TreeNode node in nodes)
                                                {
                                                    if (node.getName().Replace("$", "") == cellInRange.Address.Replace("$", "") && node.getWorksheet() == cellInRange.Worksheet.Name)
                                                    {
                                                        input_cell = node;
                                                    }
                                                    else
                                                        continue;
                                                }
                                            }
                                            //If it wasn't found, then it is blank, and we have to create a TreeNode for it
                                            if (input_cell == null)
                                            {
                                                input_cell = new TreeNode(cellInRange.Address, cellInRange.Worksheet.Name);
                                                if (toggle_array_storage.Checked)
                                                {
                                                    //Check if this cell's coordinates are within the bounds of the used range, otherwise there will be an index out of bounds error
                                                    if (cellInRange.Column <= cellInRange.Worksheet.UsedRange.Columns.Count && cellInRange.Row <= cellInRange.Worksheet.UsedRange.Rows.Count)
                                                    {
                                                        nodes_grid[cellInRange.Worksheet.Index - 1][cellInRange.Row - 1][cellInRange.Column - 1] = input_cell;
                                                    }
                                                }
                                                else  //toggle_array_storage not checked
                                                {
                                                    nodes.Add(input_cell);
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
                                    if (toggle_reporting.Checked)
                                    {
                                        MessageBox.Show("OK 3");
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
                                        if (toggle_array_storage.Checked)
                                        {
                                            //Check if this cell's coordinates are within the bounds of the used range of its spreadsheet, otherwise there will be an index out of bounds error
                                            if (input.Column <= input.Worksheet.UsedRange.Columns.Count && input.Row <= input.Worksheet.UsedRange.Rows.Count)
                                            {
                                                //if a TreeNode exists for this cell already, use it
                                                if (nodes_grid[input.Worksheet.Index - 1][input.Row - 1][input.Column - 1] != null)
                                                {
                                                    input_cell = nodes_grid[input.Worksheet.Index - 1][input.Row - 1][input.Column - 1];
                                                }
                                            }
                                        }
                                        else //toggle_array_storage not checked
                                        {
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
                                        }
                                        //If it wasn't found, then it is blank, and we have to create a TreeNode for it
                                        if (input_cell == null)
                                        {
                                            input_cell = new TreeNode(cell_coordinates.Replace("$", ""), ws_name);
                                            if (toggle_array_storage.Checked)
                                            {
                                                //Check if this cell's coordinates are within the bounds of the used range, otherwise there will be an index out of bounds error
                                                if (input.Column <= input.Worksheet.UsedRange.Columns.Count && input.Row <= input.Worksheet.UsedRange.Rows.Count)
                                                {
                                                    nodes_grid[input.Worksheet.Index - 1][input.Row - 1][input.Column - 1] = input_cell;
                                                }
                                            }
                                            else //toggle_array_storage not checked
                                            {
                                                nodes.Add(input_cell);
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
                                    if (toggle_reporting.Checked)
                                    {
                                        MessageBox.Show("OK 4");
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
                                        if (toggle_array_storage.Checked)
                                        {
                                            //Check if this cell's coordinates are within the bounds of the used range of its spreadsheet, otherwise there will be an index out of bounds error
                                            if (input.Column <= input.Worksheet.UsedRange.Columns.Count && input.Row <= input.Worksheet.UsedRange.Rows.Count)
                                            {
                                                //if a TreeNode exists for this cell already, use it
                                                if (nodes_grid[input.Worksheet.Index - 1][input.Row - 1][input.Column - 1] != null)
                                                {
                                                    input_cell = nodes_grid[input.Worksheet.Index - 1][input.Row - 1][input.Column - 1];
                                                }
                                            }
                                        }
                                        else //toggle_array_storage not checked
                                        {
                                            foreach (TreeNode node in nodes)
                                            {
                                                if (node.getName().Replace("$", "") == cell_coordinates.Replace("$", "") && node.getWorksheet().Equals(ws_name))
                                                {
                                                    input_cell = node;
                                                }
                                                else
                                                {
                                                    continue;
                                                }
                                            }
                                        }
                                        //If it wasn't found, then it is blank, and we have to create a TreeNode for it
                                        if (input_cell == null)
                                        {
                                            input_cell = new TreeNode(cell_coordinates.Replace("$", ""), ws_name);
                                            if (toggle_array_storage.Checked)
                                            {
                                                //Check if this cell's coordinates are within the bounds of the used range, otherwise there will be an index out of bounds error
                                                if (input.Column <= input.Worksheet.UsedRange.Columns.Count && input.Row <= input.Worksheet.UsedRange.Rows.Count)
                                                {
                                                    nodes_grid[input.Worksheet.Index - 1][input.Row - 1][input.Column - 1] = input_cell;
                                                }
                                            }
                                            else //toggle_array_storage not checked
                                            {
                                                nodes.Add(input_cell);
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
                                if (toggle_reporting.Checked)
                                {
                                    MessageBox.Show("OK 5");
                                }
                                //List<Excel.Range> rangeList = new List<Excel.Range>();
                                foreach (Match match in matchedRanges)
                                {
                                    formula = formula.Replace(match.Value, "");
                                    string[] endCells = match.Value.Split(':');     //Split up each matched range into the start and end cells of the range
                                    TreeNode range = null;
                                    //Try to find the range in existing TreeNodes
                                    if (toggle_array_storage.Checked)
                                    {
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
                                    }
                                    else //toggle_array_storage not checked
                                    {
                                        foreach (TreeNode n in nodes)
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
                                    }
                                    //If it does not exist, create it
                                    if (range == null)
                                    {
                                        //MessageBox.Show("Created range node:" + c.Worksheet.Name + "_" + endCells[0] + ":" + endCells[1]);
                                        range = new TreeNode(endCells[0].Replace("$", "") + "_to_" + endCells[1].Replace("$", ""), c.Worksheet.Name);
                                        if (toggle_array_storage.Checked)
                                        {
                                            ranges.Add(range);
                                        }
                                        else
                                        {
                                            nodes.Add(range);
                                        }
                                    }
                                    formula_cell.addParent(range);
                                    range.addChild(formula_cell);
                                    //Add each cell contained in the range to the dependencies
                                    foreach (Excel.Range cellInRange in c.Worksheet.Range[endCells[0], endCells[1]])
                                    {
                                        TreeNode input_cell = null;
                                        //Find the node object for the current cell in the existing TreeNodes
                                        //HERE HERE
                                        if (toggle_array_storage.Checked)
                                        {
                                            //Check if this cell's coordinates are within the bounds of the used range, otherwise there will be an index out of bounds error
                                            if (cellInRange.Column <= cellInRange.Worksheet.UsedRange.Columns.Count && cellInRange.Row <= cellInRange.Worksheet.UsedRange.Rows.Count)
                                            {
                                                //if a TreeNode exists for this cell already, use it
                                                if (nodes_grid[cellInRange.Worksheet.Index - 1][cellInRange.Row - 1][cellInRange.Column - 1] != null)
                                                {
                                                    input_cell = nodes_grid[cellInRange.Worksheet.Index - 1][cellInRange.Row - 1][cellInRange.Column - 1];
                                                }
                                            }
                                        }
                                        else //toggle_array_storage not checked
                                        {
                                            foreach (TreeNode node in nodes)
                                            {
                                                if (node.getName().Replace("$", "") == cellInRange.Address.Replace("$", "") && node.getWorksheet() == cellInRange.Worksheet.Name)
                                                {
                                                    input_cell = node;
                                                }
                                                else
                                                {
                                                    continue;
                                                }
                                            }
                                        }
                                        //If it wasn't found, then it is blank, and we have to create a TreeNode for it
                                        if (input_cell == null)
                                        {
                                            input_cell = new TreeNode(cellInRange.Address, cellInRange.Worksheet.Name);
                                            if (toggle_array_storage.Checked)
                                            {
                                                //Check if this cell's coordinates are within the bounds of the used range, otherwise there will be an index out of bounds error
                                                if (cellInRange.Column <= cellInRange.Worksheet.UsedRange.Columns.Count && cellInRange.Row <= cellInRange.Worksheet.UsedRange.Rows.Count)
                                                {
                                                    nodes_grid[cellInRange.Worksheet.Index - 1][cellInRange.Row - 1][cellInRange.Column - 1] = input_cell;
                                                }
                                            }
                                            else //toggle_array_storage not checked
                                            {
                                                nodes.Add(input_cell);
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
                                    
                                    string[] endCells = named_range.RefersToRange.Address.Split(':');     //Split up each named range into the start and end cells of the range
                                    TreeNode range = null;
                                    //Try to find the range in existing TreeNodes
                                    if (toggle_array_storage.Checked)
                                    {
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
                                    }
                                    else //toggle_array_storage not checked
                                    {
                                        foreach (TreeNode n in nodes)
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
                                    }
                                    //If it does not exist, create it
                                    if (range == null)
                                    {
                                        //MessageBox.Show("Created range node:" + c.Worksheet.Name + "_" + endCells[0] + ":" + endCells[1]);
                                        range = new TreeNode(endCells[0].Replace("$", "") + "_to_" + endCells[1].Replace("$", ""), named_range.RefersToRange.Worksheet.Name);
                                        if (toggle_array_storage.Checked)
                                        {
                                            ranges.Add(range);
                                        }
                                        else
                                        {
                                            nodes.Add(range);
                                        }
                                    }
                                    formula_cell.addParent(range);
                                    range.addChild(formula_cell);
                                    //Add each cell contained in the range to the dependencies
                                    foreach (Excel.Range cellInRange in named_range.RefersToRange.Worksheet.Range[endCells[0], endCells[1]])
                                    {
                                        TreeNode input_cell = null;
                                        //Find the node object for the current cell in the existing TreeNodes
                                        if (toggle_array_storage.Checked)
                                        {
                                            //Check if this cell's coordinates are within the bounds of the used range, otherwise there will be an index out of bounds error
                                            if (cellInRange.Column <= cellInRange.Worksheet.UsedRange.Columns.Count && cellInRange.Row <= cellInRange.Worksheet.UsedRange.Rows.Count)
                                            {
                                                //if a TreeNode exists for this cell already, use it
                                                if (nodes_grid[cellInRange.Worksheet.Index - 1][cellInRange.Row - 1][cellInRange.Column - 1] != null)
                                                {
                                                    input_cell = nodes_grid[cellInRange.Worksheet.Index - 1][cellInRange.Row - 1][cellInRange.Column - 1];
                                                }
                                            }
                                        }
                                        else //toggle_array_storage not checked
                                        {
                                            foreach (TreeNode node in nodes)
                                            {
                                                if (node.getName().Replace("$", "") == cellInRange.Address.Replace("$", "") && node.getWorksheet() == cellInRange.Worksheet.Name)
                                                {
                                                    input_cell = node;
                                                }
                                                else
                                                    continue;
                                            }
                                        }
                                        //If it wasn't found, then it is blank, and we have to create a TreeNode for it
                                        if (input_cell == null)
                                        {
                                            input_cell = new TreeNode(cellInRange.Address, cellInRange.Worksheet.Name);
                                            if (toggle_array_storage.Checked)
                                            {
                                                //Check if this cell's coordinates are within the bounds of the used range, otherwise there will be an index out of bounds error
                                                if (cellInRange.Column <= cellInRange.Worksheet.UsedRange.Columns.Count && cellInRange.Row <= cellInRange.Worksheet.UsedRange.Rows.Count)
                                                {
                                                    nodes_grid[cellInRange.Worksheet.Index - 1][cellInRange.Row - 1][cellInRange.Column - 1] = input_cell;
                                                }
                                            }
                                            else //toggle_array_storage not checked
                                            {
                                                nodes.Add(input_cell);
                                            }
                                        }

                                        //Update the dependencies
                                        range.addParent(input_cell);
                                        input_cell.addChild(range);
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
                                if (toggle_reporting.Checked)
                                {
                                    MessageBox.Show("OK 6");
                                }
                                foreach (Match m in matchedCells)
                                {
                                    Excel.Range input = c.Worksheet.get_Range(m.Value);
                                    TreeNode input_cell = null;
                                    //MessageBox.Show(m.Value);
                                    //Find the node object for the current cell in the existing TreeNodes
                                    if (toggle_array_storage.Checked)
                                    {
                                        //Check if this cell's coordinates are within the bounds of the used range, otherwise there will be an index out of bounds error
                                        if (input.Column <= input.Worksheet.UsedRange.Columns.Count && input.Row <= input.Worksheet.UsedRange.Rows.Count)
                                        {
                                            //if a TreeNode exists for this cell already, use it
                                            if (nodes_grid[input.Worksheet.Index - 1][input.Row - 1][input.Column - 1] != null)
                                            {
                                                input_cell = nodes_grid[input.Worksheet.Index - 1][input.Row - 1][input.Column - 1];
                                            }
                                        }
                                    }
                                    else //toggle_array_storage not checked
                                    {
                                        foreach (TreeNode node in nodes)
                                        {
                                            if (node.getName().Replace("$", "") == m.Value.Replace("$", "") && node.getWorksheet() == c.Worksheet.Name)
                                            {
                                                input_cell = node;
                                            }
                                            else
                                            {
                                                continue;
                                            }
                                        }
                                    }
                                    //If it wasn't found, then it is blank, and we have to create a TreeNode for it
                                    if (input_cell == null)
                                    {
                                        input_cell = new TreeNode(m.Value.Replace("$", ""), c.Worksheet.Name);
                                        if (toggle_array_storage.Checked)
                                        {
                                            //Check if this cell's coordinates are within the bounds of the used range, otherwise there will be an index out of bounds error
                                            if (input.Column <= input.Worksheet.UsedRange.Columns.Count && input.Row <= input.Worksheet.UsedRange.Rows.Count)
                                            {
                                                nodes_grid[input.Worksheet.Index - 1][input.Row - 1][input.Column - 1] = input_cell;
                                            }
                                        }
                                        else //toggle_array_storage not checked
                                        {
                                            nodes.Add(input_cell);
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
                TreeNode chart_node = new TreeNode(chart.Name, "none");
                chart_node.setChart(true);
                if (toggle_array_storage.Checked)
                {
                    charts.Add(chart_node);
                }
                else //toggle_array_storage not checked
                {
                    nodes.Add(chart_node);
                }
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
                            if (toggle_array_storage.Checked)
                            {
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
                            }
                            else //toggle_array_storage not checked
                            {
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
                            }
                            //If it was not found, create it
                            if (range == null)
                            {
                                range = new TreeNode(endCells[0].Replace("$", "") + "_to_" + endCells[1].Replace("$", ""), ws_name);
                                //MessageBox.Show("Created range node:" + ws_name + "_" + endCells[0] + ":" + endCells[1]);
                                if (toggle_array_storage.Checked)
                                {
                                    ranges.Add(range);
                                }
                                else //toggle_array_storage not checked
                                {
                                    nodes.Add(range);
                                }
                            }
                            chart_node.addParent(range);
                            range.addChild(chart_node);
                            //Add each cell contained in the range to the dependencies
                            foreach (Excel.Range cellInRange in Globals.ThisAddIn.Application.Worksheets[ws_index].Range[endCells[0], endCells[1]])
                            {
                                TreeNode input_cell = null;
                                //Find the node object for the current cell in the existing TreeNodes
                                if (toggle_array_storage.Checked)
                                {
                                    //Check if this cell's coordinates are within the bounds of the used range, otherwise there will be an index out of bounds error
                                    if (cellInRange.Column <= cellInRange.Worksheet.UsedRange.Columns.Count && cellInRange.Row <= cellInRange.Worksheet.UsedRange.Rows.Count)
                                    {
                                        //if a TreeNode exists for this cell already
                                        if (nodes_grid[cellInRange.Worksheet.Index - 1][cellInRange.Row - 1][cellInRange.Column - 1] != null)
                                        {
                                            input_cell = nodes_grid[cellInRange.Worksheet.Index - 1][cellInRange.Row - 1][cellInRange.Column - 1];
                                        }
                                    }
                                }
                                else  //toggle_array_storage not checked
                                {
                                    foreach (TreeNode node in nodes)
                                    {
                                        if (node.getName().Replace("$", "") == cellInRange.Address.Replace("$", "") && node.getWorksheet() == cellInRange.Worksheet.Name)
                                        {
                                            input_cell = node;
                                        }
                                        else
                                            continue;
                                    }
                                }
                                //If it wasn't found, then it is blank, and we have to create a TreeNode for it
                                if (input_cell == null)
                                {
                                    input_cell = new TreeNode(cellInRange.Address, cellInRange.Worksheet.Name);
                                    if (toggle_array_storage.Checked)
                                    {
                                        //Check if this cell's coordinates are within the bounds of the used range, otherwise there will be an index out of bounds error
                                        if (cellInRange.Column <= cellInRange.Worksheet.UsedRange.Columns.Count && cellInRange.Row <= cellInRange.Worksheet.UsedRange.Rows.Count)
                                        {
                                            nodes_grid[cellInRange.Worksheet.Index - 1][cellInRange.Row - 1][cellInRange.Column - 1] = input_cell;
                                        }
                                    }
                                    else  //toggle_array_storage not checked
                                    {
                                        nodes.Add(input_cell);
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
                            if (toggle_array_storage.Checked)
                            {
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
                            }
                            else //toggle_array_storage not checked
                            {
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
                            }
                            //If it was not found, create it
                            if (range == null)
                            {
                                range = new TreeNode(endCells[0].Replace("$", "") + "_to_" + endCells[1].Replace("$", ""), ws_name);
                                //MessageBox.Show("Created range node:" + ws_name + "_" + endCells[0] + ":" + endCells[1]);
                                if (toggle_array_storage.Checked)
                                {
                                    ranges.Add(range);
                                }
                                else //toggle_array_storage not checked
                                {
                                    nodes.Add(range);
                                }
                            }
                            chart_node.addParent(range);
                            range.addChild(chart_node);
                            //Add each cell contained in the range to the dependencies
                            foreach (Excel.Range cellInRange in Globals.ThisAddIn.Application.Worksheets[ws_index].Range[endCells[0], endCells[1]])
                            {
                                TreeNode input_cell = null;
                                //Find the node object for the current cell in the existing TreeNodes
                                if (toggle_array_storage.Checked)
                                {
                                    //Check if this cell's coordinates are within the bounds of the used range, otherwise there will be an index out of bounds error
                                    if (cellInRange.Column <= cellInRange.Worksheet.UsedRange.Columns.Count && cellInRange.Row <= cellInRange.Worksheet.UsedRange.Rows.Count)
                                    {
                                        //if a TreeNode exists for this cell already
                                        if (nodes_grid[cellInRange.Worksheet.Index - 1][cellInRange.Row - 1][cellInRange.Column - 1] != null)
                                        {
                                            input_cell = nodes_grid[cellInRange.Worksheet.Index - 1][cellInRange.Row - 1][cellInRange.Column - 1];
                                        }
                                    }
                                }
                                else  //toggle_array_storage not checked
                                {
                                    foreach (TreeNode node in nodes)
                                    {
                                        if (node.getName().Replace("$", "") == cellInRange.Address.Replace("$", "") && node.getWorksheet() == cellInRange.Worksheet.Name)
                                        {
                                            input_cell = node;
                                        }
                                        else
                                            continue;
                                    }
                                }
                                //If it wasn't found, then it is blank, and we have to create a TreeNode for it
                                if (input_cell == null)
                                {
                                    input_cell = new TreeNode(cellInRange.Address, cellInRange.Worksheet.Name);
                                    if (toggle_array_storage.Checked)
                                    {
                                        //Check if this cell's coordinates are within the bounds of the used range, otherwise there will be an index out of bounds error
                                        if (cellInRange.Column <= cellInRange.Worksheet.UsedRange.Columns.Count && cellInRange.Row <= cellInRange.Worksheet.UsedRange.Rows.Count)
                                        {
                                            nodes_grid[cellInRange.Worksheet.Index - 1][cellInRange.Row - 1][cellInRange.Column - 1] = input_cell;
                                        }
                                    }
                                    else  //toggle_array_storage not checked
                                    {
                                        nodes.Add(input_cell);
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
                            if (toggle_array_storage.Checked)
                            {
                                //Check if this cell's coordinates are within the bounds of the used range of its spreadsheet, otherwise there will be an index out of bounds error
                                if (input.Column <= input.Worksheet.UsedRange.Columns.Count && input.Row <= input.Worksheet.UsedRange.Rows.Count)
                                {
                                    //if a TreeNode exists for this cell already, use it
                                    if (nodes_grid[input.Worksheet.Index - 1][input.Row - 1][input.Column - 1] != null)
                                    {
                                        input_cell = nodes_grid[input.Worksheet.Index - 1][input.Row - 1][input.Column - 1];
                                    }
                                }
                            }
                            else //toggle_array_storage not checked
                            {
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
                            }
                            //If it wasn't found, then it is blank, and we have to create a TreeNode for it
                            if (input_cell == null)
                            {
                                input_cell = new TreeNode(cell_coordinates.Replace("$", ""), ws_name);
                                if (toggle_array_storage.Checked)
                                {
                                    //Check if this cell's coordinates are within the bounds of the used range, otherwise there will be an index out of bounds error
                                    if (input.Column <= input.Worksheet.UsedRange.Columns.Count && input.Row <= input.Worksheet.UsedRange.Rows.Count)
                                    {
                                        nodes_grid[input.Worksheet.Index - 1][input.Row - 1][input.Column - 1] = input_cell;
                                    }
                                }
                                else //toggle_array_storage not checked
                                {
                                    nodes.Add(input_cell);
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
                            if (toggle_array_storage.Checked)
                            {
                                //Check if this cell's coordinates are within the bounds of the used range of its spreadsheet, otherwise there will be an index out of bounds error
                                if (input.Column <= input.Worksheet.UsedRange.Columns.Count && input.Row <= input.Worksheet.UsedRange.Rows.Count)
                                {
                                    //if a TreeNode exists for this cell already, use it
                                    if (nodes_grid[input.Worksheet.Index - 1][input.Row - 1][input.Column - 1] != null)
                                    {
                                        input_cell = nodes_grid[input.Worksheet.Index - 1][input.Row - 1][input.Column - 1];
                                    }
                                }
                            }
                            else //toggle_array_storage not checked
                            {
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
                            }
                            //If it wasn't found, then it is blank, and we have to create a TreeNode for it
                            if (input_cell == null)
                            {
                                input_cell = new TreeNode(cell_coordinates.Replace("$", ""), ws_name);
                                if (toggle_array_storage.Checked)
                                {
                                    //Check if this cell's coordinates are within the bounds of the used range, otherwise there will be an index out of bounds error
                                    if (input.Column <= input.Worksheet.UsedRange.Columns.Count && input.Row <= input.Worksheet.UsedRange.Rows.Count)
                                    {
                                        nodes_grid[input.Worksheet.Index - 1][input.Row - 1][input.Column - 1] = input_cell;
                                    }
                                }
                                else //toggle_array_storage not checked
                                {
                                    nodes.Add(input_cell);
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
                        if (toggle_array_storage.Checked)
                        {
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
                        }
                        else //toggle_array_storage not checked
                        {
                            foreach (TreeNode n in nodes)
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
                        }
                        //If it does not exist, create it
                        if (range == null)
                        {
                            //MessageBox.Show("Created range node:" + c.Worksheet.Name + "_" + endCells[0] + ":" + endCells[1]);
                            range = new TreeNode(endCells[0].Replace("$", "") + "_to_" + endCells[1].Replace("$", ""), named_range.RefersToRange.Worksheet.Name);
                            if (toggle_array_storage.Checked)
                            {
                                ranges.Add(range);
                            }
                            else
                            {
                                nodes.Add(range);
                            }
                        }
                        //Update dependencies
                        chart_node.addParent(range);
                        range.addChild(chart_node);
                        //Add each cell contained in the range to the dependencies
                        foreach (Excel.Range cellInRange in named_range.RefersToRange.Worksheet.Range[endCells[0], endCells[1]])
                        {
                            TreeNode input_cell = null;
                            //Find the node object for the current cell in the existing TreeNodes
                            if (toggle_array_storage.Checked)
                            {
                                //Check if this cell's coordinates are within the bounds of the used range, otherwise there will be an index out of bounds error
                                if (cellInRange.Column <= cellInRange.Worksheet.UsedRange.Columns.Count && cellInRange.Row <= cellInRange.Worksheet.UsedRange.Rows.Count)
                                {
                                    //if a TreeNode exists for this cell already, use it
                                    if (nodes_grid[cellInRange.Worksheet.Index - 1][cellInRange.Row - 1][cellInRange.Column - 1] != null)
                                    {
                                        input_cell = nodes_grid[cellInRange.Worksheet.Index - 1][cellInRange.Row - 1][cellInRange.Column - 1];
                                    }
                                }
                            }
                            else //toggle_array_storage not checked
                            {
                                foreach (TreeNode node in nodes)
                                {
                                    if (node.getName().Replace("$", "") == cellInRange.Address.Replace("$", "") && node.getWorksheet() == cellInRange.Worksheet.Name)
                                    {
                                        input_cell = node;
                                    }
                                    else
                                        continue;
                                }
                            }
                            //If it wasn't found, then it is blank, and we have to create a TreeNode for it
                            if (input_cell == null)
                            {
                                input_cell = new TreeNode(cellInRange.Address, cellInRange.Worksheet.Name);
                                if (toggle_array_storage.Checked)
                                {
                                    //Check if this cell's coordinates are within the bounds of the used range, otherwise there will be an index out of bounds error
                                    if (cellInRange.Column <= cellInRange.Worksheet.UsedRange.Columns.Count && cellInRange.Row <= cellInRange.Worksheet.UsedRange.Rows.Count)
                                    {
                                        nodes_grid[cellInRange.Worksheet.Index - 1][cellInRange.Row - 1][cellInRange.Column - 1] = input_cell;
                                    }
                                }
                                else //toggle_array_storage not checked
                                {
                                    nodes.Add(input_cell);
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
                    TreeNode chart_node = new TreeNode(chart.Name, worksheet.Name);
                    chart_node.setChart(true);
                    nodes.Add(chart_node);
                    foreach (Excel.Series series in (Excel.SeriesCollection)chart.Chart.SeriesCollection(Type.Missing))
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
                                    range = new TreeNode(endCells[0].Replace("$", "") + "_to_" + endCells[1].Replace("$", ""), ws_name);
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
                                        input_cell = new TreeNode(cellInRange.Address, cellInRange.Worksheet.Name);
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
                                    range = new TreeNode(endCells[0].Replace("$", "") + "_to_" + endCells[1].Replace("$", ""), ws_name);
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
                                        input_cell = new TreeNode(cellInRange.Address, cellInRange.Worksheet.Name);
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
                                    input_cell = new TreeNode(cell_coordinates.Replace("$", ""), ws_name);
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
                                    input_cell = new TreeNode(cell_coordinates.Replace("$", ""), ws_name);
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
                                range = new TreeNode(endCells[0].Replace("$", "") + "_to_" + endCells[1].Replace("$", ""), worksheet.Name);
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
                                    input_cell = new TreeNode(cellInRange.Address, cellInRange.Worksheet.Name);
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
                            if (toggle_array_storage.Checked)
                            {
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
                            }
                            else //toggle_array_storage not checked
                            {
                                foreach (TreeNode n in nodes)
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
                            }
                            //If it does not exist, create it
                            if (range == null)
                            {
                                //MessageBox.Show("Created range node:" + c.Worksheet.Name + "_" + endCells[0] + ":" + endCells[1]);
                                range = new TreeNode(endCells[0].Replace("$", "") + "_to_" + endCells[1].Replace("$", ""), named_range.RefersToRange.Worksheet.Name);
                                if (toggle_array_storage.Checked)
                                {
                                    ranges.Add(range);
                                }
                                else
                                {
                                    nodes.Add(range);
                                }
                            }
                            //Update dependencies
                            chart_node.addParent(range);
                            range.addChild(chart_node);
                            //Add each cell contained in the range to the dependencies
                            foreach (Excel.Range cellInRange in named_range.RefersToRange.Worksheet.Range[endCells[0], endCells[1]])
                            {
                                TreeNode input_cell = null;
                                //Find the node object for the current cell in the existing TreeNodes
                                if (toggle_array_storage.Checked)
                                {
                                    //Check if this cell's coordinates are within the bounds of the used range, otherwise there will be an index out of bounds error
                                    if (cellInRange.Column <= cellInRange.Worksheet.UsedRange.Columns.Count && cellInRange.Row <= cellInRange.Worksheet.UsedRange.Rows.Count)
                                    {
                                        //if a TreeNode exists for this cell already, use it
                                        if (nodes_grid[cellInRange.Worksheet.Index - 1][cellInRange.Row - 1][cellInRange.Column - 1] != null)
                                        {
                                            input_cell = nodes_grid[cellInRange.Worksheet.Index - 1][cellInRange.Row - 1][cellInRange.Column - 1];
                                        }
                                    }
                                }
                                else //toggle_array_storage not checked
                                {
                                    foreach (TreeNode node in nodes)
                                    {
                                        if (node.getName().Replace("$", "") == cellInRange.Address.Replace("$", "") && node.getWorksheet() == cellInRange.Worksheet.Name)
                                        {
                                            input_cell = node;
                                        }
                                        else
                                            continue;
                                    }
                                }
                                //If it wasn't found, then it is blank, and we have to create a TreeNode for it
                                if (input_cell == null)
                                {
                                    input_cell = new TreeNode(cellInRange.Address, cellInRange.Worksheet.Name);
                                    if (toggle_array_storage.Checked)
                                    {
                                        //Check if this cell's coordinates are within the bounds of the used range, otherwise there will be an index out of bounds error
                                        if (cellInRange.Column <= cellInRange.Worksheet.UsedRange.Columns.Count && cellInRange.Row <= cellInRange.Worksheet.UsedRange.Rows.Count)
                                        {
                                            nodes_grid[cellInRange.Worksheet.Index - 1][cellInRange.Row - 1][cellInRange.Column - 1] = input_cell;
                                        }
                                    }
                                    else //toggle_array_storage not checked
                                    {
                                        nodes.Add(input_cell);
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
                                input_cell = new TreeNode(m.Value, worksheet.Name);
                                nodes.Add(input_cell);
                            }

                            //Update the dependencies
                            chart_node.addParent(input_cell);
                            input_cell.addChild(chart_node);
                        }

                    }
                }
            }
            //Propagate weights  -- static propagation in the dependence graph (no swapping of values)
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
                                if (!node.hasParents()) //hasChildren())
                                {
                                    node.setWeight(1.0);  //Set the weight of all output nodes (and charts) to 1.0 to start
                                    //Now we propagate proportional weights to all of this node's inputs
                                    //propagateWeight(node, 1.0);
                                    propagateWeightUp(node, 1.0);
                                }
                            }
                        }
                    }
                }
            }
            
            //double max_weight = 0.0;  //Keep track of the max weight for normalizing later (used for coloring cells based on weight)
            //foreach (TreeNode node in nodes)
            //{
            //    if (node.getWeight() > max_weight)
            //        max_weight = node.getWeight();
            //}
            //TODO -- we are not able to capture ranges that are identified in stored procedures or macros, just ones referenced in formulas

            //TODO -- Dealing with fuzzing of charts -- idea: any cell that feeds into a chart is essentially an output; the chart is just a visual representation (can charts operate on values before they are displayed? don't think so...)

            List<StartValue> starting_outputs = new List<StartValue>(); //This will store the values of all the output nodes at the start of the procedure for swapping values (fuzzing)
            List<TreeNode> output_cells = new List<TreeNode>(); //This will store all the output nodes at the start of the fuzzing procedure
            //Store all the starting output values
            if (toggle_array_storage.Checked)
            {
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
            }
            else //toggle_array_storage not checked
            {
                foreach (TreeNode node in nodes)
                {
                    //if (!node.hasChildren() && !node.isChart()) //If the node does not feed into any other nodes, and it is not a chart, then it is considered output
                    if (!node.hasChildren() && node.hasParents()) //Nodes that do not feed into any other nodes are considered output, unless nothing feeds into them either. 
                    {
                        output_cells.Add(node);
                    }
                    /**
                    //We also want to add any nodes that feed into charts, because they're essentially outputs. The chart is just a visual aid. 
                    //Nodes feeding into a chart will either be cell nodes or range nodes; for ranges, we should add every cell in the range to output_cells
                    //We also need to make sure we are not adding duplicates in this case
                    if (node.isChart())
                    {
                        List<TreeNode> chart_parents = node.getParents();
                        foreach (TreeNode chart_parent in chart_parents)
                        {
                            if (!chart_parent.isRange()) //If it is a single cell node
                            {
                                //Check for duplicate entries
                                bool parent_already_added = false;
                                foreach (TreeNode output_cell in output_cells)
                                {
                                    if (chart_parent.getName() == output_cell.getName())
                                        parent_already_added = true;
                                }
                                //If the chart parent is not on the list, add it
                                if (!parent_already_added)
                                    output_cells.Add(chart_parent);
                            }
                            else if (chart_parent.isRange())
                            {
                                List<TreeNode> range_parents = chart_parent.getParents();
                                foreach (TreeNode range_parent in range_parents)
                                {
                                    //Check for duplicate entries
                                    bool parent_already_added = false;
                                    foreach (TreeNode output_cell in output_cells)
                                    {
                                        if (range_parent.getName() == output_cell.getName())
                                            parent_already_added = true;
                                    }
                                    //If the range parent is not on the list, add it
                                    if (!parent_already_added)
                                        output_cells.Add(range_parent);
                                }
                            }

                        }
                    }
                    **/
                }
            }

            //This part stores all the output values before any perturbations are applied
            foreach (TreeNode n in output_cells)
            {
                // If the TreeNode is a chart
                if (n.isChart())
                {
                    // Add a StartValue with the average of the range of inputs for each range of inputs
                    //MessageBox.Show(n.getName() + " is output.");
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
                    //n.setOriginalColor(System.Drawing.ColorTranslator.FromOle((int)cell.Interior.Color)); 
                    //Only store the original color on the first run of the tool
                    //if (toolHasNotRun == false)
                    //{
                        //n.setOriginalColor(cell.Interior.ColorIndex);
                    //    n.setOriginalColor(System.Drawing.ColorTranslator.FromOle((int)cell.Interior.Color));
                    //}
                    //Color final outputs of computation in red:
                    //cell.Interior.Color = System.Drawing.Color.Blue;
                    //MessageBox.Show(cell.Address + " is an output cell.");
                    try     //If the output is a number
                    {
                        double d = (double)nodeWorksheet.get_Range(n.getName()).Value;
                        StartValue sv = new StartValue(d);
                        starting_outputs.Add(sv); //Try adding it as a number
                    }
                    catch   //If the output is a string
                    {
                        string s = nodeWorksheet.get_Range(n.getName()).Value;
                        StartValue sv = new StartValue(s);
                        starting_outputs.Add(sv); //starting_outputs.Add(activeWorksheet.get_Range(n.getName()).Value); //If it doesn't work, it must be a string output
                    }
                }
            }
            swTree.Stop();
            // Get the elapsed time from tsTree as a TimeSpan value.
            TimeSpan tsTree = swTree.Elapsed;
            // Format and display the TimeSpan value. 
            string treeTime = String.Format("{0:00}:{1:00}:{2:00}.{3:00}", tsTree.Hours, tsTree.Minutes, tsTree.Seconds, tsTree.Milliseconds / 10);
            //MessageBox.Show("Done building dependence graph.\nTime elapsed: " + treeTime);
            
            System.Diagnostics.Stopwatch sw = new System.Diagnostics.Stopwatch();
            sw.Start();

            //Grids for storing influences
            double[][][] influences_grid = null;
            int[][][] times_perturbed = null;
            //if (toggle_global_perturbation.Checked)
            //{
                influences_grid = new double[Globals.ThisAddIn.Application.Worksheets.Count + Globals.ThisAddIn.Application.Charts.Count][][];
                times_perturbed = new int[Globals.ThisAddIn.Application.Worksheets.Count + Globals.ThisAddIn.Application.Charts.Count][][];
                foreach (Excel.Worksheet worksheet in Globals.ThisAddIn.Application.Worksheets)
                {
                    influences_grid[worksheet.Index - 1] = new double[worksheet.UsedRange.Rows.Count][];
                    times_perturbed[worksheet.Index - 1] = new int[worksheet.UsedRange.Rows.Count][];
                    for (int row = 0; row < worksheet.UsedRange.Rows.Count; row++)
                    {
                        influences_grid[worksheet.Index - 1][row] = new double[worksheet.UsedRange.Columns.Count];
                        times_perturbed[worksheet.Index - 1][row] = new int[worksheet.UsedRange.Columns.Count];
                        for (int col = 0; col < worksheet.UsedRange.Columns.Count; col++)
                        {
                            influences_grid[worksheet.Index - 1][row][col] = 0.0;
                            times_perturbed[worksheet.Index - 1][row][col] = 0;
                        }
                    }
                }
            //}
            if (toggle_no_sceen_updating.Checked)
            {
                Globals.ThisAddIn.Application.ScreenUpdating = false;
            }
            //Procedure for swapping values within ranges, one cell at a time
            if (!checkBox2.Checked) //Checks if the option for swapping values simultaneously is checked
            {
                List<TreeNode> swap_domain;
                if (toggle_array_storage.Checked)  //if array storage is checked, range nodes are stored in the 'ranges' list, so those are the ones we will perturb.
                {
                    swap_domain = ranges;
                }
                else
                {
                    swap_domain = nodes;
                }

                //Initialize min_max_delta_outputs
                min_max_delta_outputs = new double[output_cells.Count][];
                for (int i = 0; i < output_cells.Count; i++)
                {
                    min_max_delta_outputs[i] = new double[2];
                    min_max_delta_outputs[i][0] = -1.0;
                    min_max_delta_outputs[i][1] = 0.0;
                }

                //Initialize impacts_grid 
                impacts_grid = new double[Globals.ThisAddIn.Application.Worksheets.Count][][][];
                foreach (Excel.Worksheet worksheet in Globals.ThisAddIn.Application.Worksheets)
                {
                    //MessageBox.Show("Dimensions: " + Globals.ThisAddIn.Application.Worksheets.Count + " x " +
                    //worksheet.UsedRange.Rows.Count + " x " +
                    //worksheet.UsedRange.Columns.Count + " x " +
                    //output_cells.Count);
                    impacts_grid[worksheet.Index - 1] = new double[worksheet.UsedRange.Rows.Count][][];
                    for (int row = 0; row < worksheet.UsedRange.Rows.Count; row++)
                    {
                        impacts_grid[worksheet.Index - 1][row] = new double[worksheet.UsedRange.Columns.Count][];
                        for (int col = 0; col < worksheet.UsedRange.Columns.Count; col++)
                        {
                            impacts_grid[worksheet.Index - 1][row][col] = new double[output_cells.Count];
                            //MessageBox.Show("output cells count = " + output_cells.Count);
                            for (int i = 0; i < output_cells.Count; i++)
                            {
                                impacts_grid[worksheet.Index - 1][row][col][i] = 0.0;
                            }
                        }
                    }
                }

                foreach (TreeNode node in swap_domain)
                {
                    //If this node is not a range, continue to the next node because no perturbations can be done on this node.
                    if (!node.isRange())
                    {
                        continue;
                    }
                    //bool all_children_are_charts = true;  //We need to know if all children are charts because we can't compute a delta for a chart
                    //if(node.isRange() && node.hasChildren())
                    //{
                    //    foreach (TreeNode child in node.getChildren())
                    //    {
                    //        if (!child.isChart())
                    //        {
                    //            all_children_are_charts = false;
                    //            break; //Do not need to continue looping because all_children_are_charts was already set to false
                    //        }
                    //    }
                    //}
                    //For every range node
                    double[] influences = new double[node.getParents().Count]; //Array to keep track of the influence values for every cell in the range
                    int influence_index = 0;        //Keeps track of the current position in the influences array
                    //double max_total_delta = 0;     //The maximum influence found (for normalizing)
                    //double min_total_delta = -1;     //The minimum influence found (for normalizing)
                    double swaps_per_range = 30.0;
                    //Swapping values; loop over all nodes in the range
                    foreach (TreeNode parent in node.getParents())
                    {
                        if (parent.hasParents()) //Do not perturb nodes which are intermediate computations
                        {
                            continue;
                        }
                        Excel.Range cell = parent.getWorksheetObject().get_Range(parent.getName());
                        string formula = "";
                        if (cell.HasFormula)
                        {
                            //MessageBox.Show("Formula: " + cell.Formula);
                            formula = cell.Formula;
                        }
                        StartValue start_value = new StartValue(cell.Value); //Store the initial value of this cell before swapping
                        double total_delta = 0.0; // Stores the total change in outputs caused by this cell after swapping with every other value in the range
                        double delta = 0.0;   // Stores the change in outputs caused by a single swap
                        //Swapping loop - swap every sibling or a reduced number of siblings (approximately equal to swaps_per_range), for reduced complexity and runtime
                        int number_neighbors_replaced = 0;
                        Random rand = new Random();
                        foreach (TreeNode sibling in node.getParents())
                        {
                            if (sibling.getName() == parent.getName() && sibling.getWorksheetObject() == parent.getWorksheetObject())
                            {
                                continue; // If this is the current cell, move on to the next cell
                            }
                            if (toggle_speed.Checked)
                            {
                                if (rand.NextDouble() > (swaps_per_range / node.getParents().Count)) //only do ~swaps_per_range swaps per range
                                {
                                    continue;
                                }
                                number_neighbors_replaced++;
                                if (toggle_global_perturbation.Checked)
                                {
                                    times_perturbed[cell.Worksheet.Index - 1][cell.Row - 1][cell.Column - 1]++;
                                }
                            }
                            else
                            {
                                if (toggle_global_perturbation.Checked)
                                {
                                    times_perturbed[cell.Worksheet.Index - 1][cell.Row - 1][cell.Column - 1]++;
                                }
                            }
                            times_perturbed[cell.Worksheet.Index - 1][cell.Row - 1][cell.Column - 1]++;
                            Excel.Range sibling_cell = sibling.getWorksheetObject().get_Range(sibling.getName());
                            cell.Value = sibling_cell.Value; //This is the swap -- we assign the value of the sibling cell to the value of our cell
                            int index = 0;  //Index for knowing which output cell we are evaluating
                            delta = 0.0;
                            foreach (TreeNode n in output_cells)
                            {
                                if (starting_outputs[index].get_string() == null) // If the output is not a string
                                {
                                    if (!n.isChart())   //If the output is not a chart, it must be a number
                                    {
                                        delta = Math.Abs(starting_outputs[index].get_double() - (double)n.getWorksheetObject().get_Range(n.getName()).Value);  //Compute the absolute change caused by the swap
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
                                        delta = Math.Abs(starting_outputs[index].get_double() - average);
                                    }
                                }
                                else  // If the output is a string
                                {
                                    //MessageBox.Show("Comparing " + starting_outputs[index].get_string() + " and " + activeWorksheet.get_Range(n.getName()).Value);
                                    if (String.Equals(starting_outputs[index].get_string(), n.getWorksheetObject().get_Range(n.getName()).Value, StringComparison.Ordinal))
                                    {
                                        delta = 0.0;
                                    }
                                    else
                                    {
                                        delta = 1.0;
                                    }
                                }
                                //Add to the impact of the cell for this output
                                //MessageBox.Show("Cell R" + (cell.Row - 1) + "C" + (cell.Column - 1) + " has " + impacts_grid[cell.Worksheet.Index - 1][cell.Row - 1][cell.Column - 1][index] + "+" + delta + " impact on output " + index);
                                impacts_grid[cell.Worksheet.Index - 1][cell.Row - 1][cell.Column - 1][index] += delta;
                                //Compare the min/max values for this output to this delta
                                if (min_max_delta_outputs[index][0] == -1.0)
                                {
                                    min_max_delta_outputs[index][0] = delta;
                                }
                                else
                                {
                                    if (min_max_delta_outputs[index][0] > delta)
                                    {
                                        min_max_delta_outputs[index][0] = delta;
                                    }
                                }
                                if (min_max_delta_outputs[index][1] < delta)
                                {
                                    min_max_delta_outputs[index][1] = delta;
                                }
                                index++;
                                total_delta = total_delta + delta;
                            }
                        }
                        
                        //if (toggle_global_perturbation.Checked)
                        //{
                        //    influences_grid[cell.Worksheet.Index - 1][cell.Row - 1][cell.Column - 1] += total_delta;
                        //}

                        //if (toggle_speed.Checked)
                        //{
                        //    if (number_neighbors_replaced != 0)
                        //    {
                        //        total_delta = total_delta / number_neighbors_replaced;
                        //    }
                        //}
                        //else
                        //{
                        //    if (node.getParents().Count - 1 != 0) //The range must have had at least 2 cells in it
                        //    {
                        //        total_delta = total_delta / (node.getParents().Count - 1); //We divide by the number of swaps to get an average per-swap delta: not really necessary since we then scale things based on the max_delta and min_delta; would be useful if the max_delta and min_delta were global for all the ranges
                        //    }
                        //}
                        //MessageBox.Show(cell.get_Address() + " Total delta = " + (total_delta * 100) + "%");                       
                        //influences[influence_index] = total_delta;
                        //influence_index++;
                        //if (max_total_delta < total_delta)
                        //{
                        //    max_total_delta = total_delta;
                        //}
                        //if (min_total_delta > total_delta || min_total_delta == -1)
                        //{
                        //    min_total_delta = total_delta;
                        //}
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
                        //parent.setOriginalColor(System.Drawing.ColorTranslator.FromOle((int)cell.Interior.Color));
                        //Only store the original color on the first run of the tool
                        //if (toolHasNotRun == false)
                        //{
                            //parent.setOriginalColor(cell.Interior.ColorIndex);
                        //    parent.setOriginalColor(System.Drawing.ColorTranslator.FromOle((int)cell.Interior.Color));
                        //}
                        //cell.Interior.Color = System.Drawing.Color.Beige;
                    }
                    
                    //int ind = 0;
                    //MessageBox.Show("MIN DELTA: " + min_total_delta + "\nMAX DELTA: " + max_total_delta);
                    //Normalize the influences based on the smallest and largest influence values. This way they are > 0 and < 1.
                    //foreach (TreeNode parent in node.getParents())
                    //{
                    //    if (max_total_delta != 0)
                    //    {
                    //        if ((influences[ind] - min_total_delta) / max_total_delta > 1) //MessageBox.Show("Influence = " + influences[ind]);
                    //        {
                    //            MessageBox.Show("Error. Influence should not be greater than 1.");
                    //            MessageBox.Show("Influence = " + influences[ind]);
                    //            MessageBox.Show("(" + influences[ind] + " - " +  min_total_delta +") / " + max_total_delta);
                    //        }
                    //        influences[ind] = (influences[ind] - min_total_delta) / max_total_delta;
                    //    }
                    //    ind++;
                    //}

                    ////Color cells based on influence
                    //if (!toggle_global_perturbation.Checked)
                    //{
                    //    if (toggle_analyze_outliers.Checked)
                    //    {
                    //        string cell1 = node.getName().Substring(0, node.getName().IndexOf("_"));
                    //        string cell2 = node.getName().Substring(node.getName().LastIndexOf("_") + 1, 
                    //                                                node.getName().Length - (node.getName().LastIndexOf("_") + 1));
                    //        Excel.Range range = node.getWorksheetObject().get_Range(cell1, cell2);
                    //        MessageBox.Show("Running Peirce on range " + node.getName());
                    //        run_peirce(range);
                    //        /**
                    //        int index = 0;
                    //        //Compute average influence
                    //        double average_influence = 0.0;
                    //        double denominator = (double)node.getParents().Count;
                    //        //TODO: if there are overflow issues consider making total_influence an array of doubles (of size 100 for instance) and use each slot as a bin for parts of the sum
                    //        //each part can be divided by the denominator and then the average_influence is the sum of the entries in the array
                    //        double total_influence = 0.0;
                    //        foreach (TreeNode parent in node.getParents())
                    //        {
                    //            //MessageBox.Show("influence: " + influences[index]);
                    //            total_influence += influences[index];
                    //            index++;
                    //        }
                    //        average_influence = total_influence / denominator;
                    //        //Compute the standard deviation
                    //        double variance = 0.0;  //stores the sum of the suqared differences from the mean divided by the denominator
                    //        index = 0;
                    //        foreach (TreeNode parent in node.getParents())
                    //        {
                    //            variance += (influences[index] - average_influence) * (influences[index] - average_influence) / denominator;
                    //            index++;
                    //        }
                    //        double standard_deviation = Math.Sqrt(variance);
                    //        //Color cells that lie further than two standard deviations away from the mean
                    //        index = 0;
                    //        foreach (TreeNode parent in node.getParents())
                    //        {
                    //            Excel.Range cell = parent.getWorksheetObject().get_Range(parent.getName());
                    //            //If inluence is more than two standard deviations away from the mean, color that cell
                    //            //TODO This doesnt seem to work like it should - only showing unusually influential cells (2 st. dev away from mean) when perturbing locally
                    //            if (Math.Abs(influences[index] - average_influence) > 2 * standard_deviation)
                    //            {
                    //                cell.Interior.Color = System.Drawing.Color.FromArgb(Convert.ToInt32(255 - influences[index] * 255), 255, 255);
                    //            }
                    //            else
                    //            {
                    //                cell.Interior.Color = System.Drawing.Color.White;
                    //            }
                    //            index++;
                    //        }
                    //         **/
                    //    }
                    //    if (!toggle_analyze_outliers.Checked)
                    //    {
                    //        int indexer = 0;
                    //        foreach (TreeNode parent in node.getParents())
                    //        {
                    //            Excel.Range cell = parent.getWorksheetObject().get_Range(parent.getName());
                    //            //parent.setOriginalColor(System.Drawing.ColorTranslator.FromOle((int)cell.Interior.Color));

                    //            //Only store the original color on the first run of the tool
                    //            //if (toolHasNotRun == false)
                    //            //{
                    //            //parent.setOriginalColor(cell.Interior.ColorIndex);
                    //            //    parent.setOriginalColor(System.Drawing.ColorTranslator.FromOle((int)cell.Interior.Color));
                    //            //}
                    //            cell.Interior.Color = System.Drawing.Color.FromArgb(Convert.ToInt32(255 - influences[indexer] * 255), 255, 255);
                    //            indexer++;
                    //        }
                    //    }
                    //}
                }
                //Now normalize the entries in impacts_grid so that they reflect per-swap averages
                int inputs_count = 0; 
                //Find the number of input cells
                foreach (Excel.Worksheet worksheet in Globals.ThisAddIn.Application.Worksheets)
                {
                    for (int row = 0; row < worksheet.UsedRange.Rows.Count; row++)
                    {
                        for (int col = 0; col < worksheet.UsedRange.Columns.Count; col++)
                        {
                            for (int i = 0; i < output_cells.Count; i++)
                            {
                                if (times_perturbed[worksheet.Index - 1][row][col] != 0)
                                {
                                    inputs_count++;
                                    //MessageBox.Show("Cell R" + (row + 1) + "C" + (col + 1) + " total impact: " + impacts_grid[worksheet.Index - 1][row][col][i]);
                                }
                             //   impacts_grid[worksheet.Index - 1][row][col][i] = impacts_grid[worksheet.Index - 1][row][col][i] / times_perturbed[worksheet.Index - 1][row][col];
                             //   impacts_grid[worksheet.Index - 1][row][col][i] = impacts_grid[worksheet.Index - 1][row][col][i] - min_max_delta_outputs[i][0] / (min_max_delta_outputs[i][1] - min_max_delta_outputs[i][0]);
                            //    MessageBox.Show("Cell R" + (row + 1) + "C" + (col + 1) + ": " + impacts_grid[worksheet.Index - 1][row][col][i] +
                            //        "\nMin output: " + min_max_delta_outputs[i][0] +
                            //        "\nMax output: " + min_max_delta_outputs[i][1]);
                            }
                        }
                    }
                }

                //Now for each output, compute the z-score of the impact of each input
                for (int i = 0; i < output_cells.Count; i++)
                {
                    //Find the mean for the output
                    double output_sum = 0.0;
                    int non_zero_entries = 0;
                    foreach (Excel.Worksheet worksheet in Globals.ThisAddIn.Application.Worksheets)
                    {
                        for (int row = 0; row < worksheet.UsedRange.Rows.Count; row++)
                        {
                            for (int col = 0; col < worksheet.UsedRange.Columns.Count; col++)
                            {
                                if (times_perturbed[worksheet.Index - 1][row][col] != 0)
                                {
                                    if (impacts_grid[worksheet.Index - 1][row][col][i] != 0.0)
                                    {
                                        non_zero_entries++;
                                        output_sum += impacts_grid[worksheet.Index - 1][row][col][i];
                                    }
                                }
                            }
                        }
                    }

                    double output_average = output_sum / (double)non_zero_entries;
                    //Find the sample standard deviation for this output
                    double distance_sum_sq = 0.0;
                    foreach (Excel.Worksheet worksheet in Globals.ThisAddIn.Application.Worksheets)
                    {
                        for (int row = 0; row < worksheet.UsedRange.Rows.Count; row++)
                        {
                            for (int col = 0; col < worksheet.UsedRange.Columns.Count; col++)
                            {
                                if (times_perturbed[worksheet.Index - 1][row][col] != 0)
                                {
                                    distance_sum_sq += Math.Pow(output_average - impacts_grid[worksheet.Index - 1][row][col][i], 2);
                                }
                            }
                        }
                    }
                    double variance = distance_sum_sq / (double)non_zero_entries;
                    double std_dev = Math.Sqrt(variance);
                    
                    //Replace entries in impacts_grid with z-scores
                    foreach (Excel.Worksheet worksheet in Globals.ThisAddIn.Application.Worksheets)
                    {
                        for (int row = 0; row < worksheet.UsedRange.Rows.Count; row++)
                        {
                            for (int col = 0; col < worksheet.UsedRange.Columns.Count; col++)
                            {
                                if (times_perturbed[worksheet.Index - 1][row][col] != 0)
                                {
                                    impacts_grid[worksheet.Index - 1][row][col][i] = Math.Abs((impacts_grid[worksheet.Index - 1][row][col][i] - output_average) / std_dev);
                                }
                            }
                        }
                    }
                }

                //Now we want to average the z-score of every input and store it
                double[][][] average_z_scores = new double[Globals.ThisAddIn.Application.Worksheets.Count][][];
                foreach (Excel.Worksheet worksheet in Globals.ThisAddIn.Application.Worksheets)
                {
                    average_z_scores[worksheet.Index - 1] = new double[worksheet.UsedRange.Rows.Count][];
                    for (int row = 0; row < worksheet.UsedRange.Rows.Count; row++)
                    {
                        average_z_scores[worksheet.Index - 1][row] = new double[worksheet.UsedRange.Columns.Count];
                    }
                }
                foreach (Excel.Worksheet worksheet in Globals.ThisAddIn.Application.Worksheets)
                {
                    for (int row = 0; row < worksheet.UsedRange.Rows.Count; row++)
                    {
                        for (int col = 0; col < worksheet.UsedRange.Columns.Count; col++)
                        {
                            //If this cell has been perturbed, find it's average z-score
                            double total_z_score = 0.0;
                            double total_output_weight = 0.0;
                            if (times_perturbed[worksheet.Index - 1][row][col] != 0)
                            {
                                for (int i = 0; i < output_cells.Count; i++)
                                {
                                    //if (toggle_weighted_average.Checked)
                                    //{
                                        total_z_score += impacts_grid[worksheet.Index - 1][row][col][i] * output_cells[i].getWeight();
                                        total_output_weight += output_cells[i].getWeight();
                                    //}
                                    //else
                                    //{
                                    //    total_z_score += impacts_grid[worksheet.Index - 1][row][col][i];
                                    //}
                                }
                            }
                            //if(toggle_weighted_average.Checked) 
                            //{
                                average_z_scores[worksheet.Index - 1][row][col] = (total_z_score / total_output_weight);
                            //}
                            //else 
                            //{
                            //    average_z_scores[worksheet.Index - 1][row][col] = (total_z_score / output_cells.Count);
                            //}
                        }
                    }
                }

                if (!toggle_weighted_average.Checked)
                {
                    //Look for outliers:
                    List<int[]> outliers = new List<int[]>();
                    for (int i = 0; i < output_cells.Count; i++)
                    {
                        foreach (Excel.Worksheet worksheet in Globals.ThisAddIn.Application.Worksheets)
                        {
                            for (int row = 0; row < worksheet.UsedRange.Rows.Count; row++)
                            {
                                for (int col = 0; col < worksheet.UsedRange.Columns.Count; col++)
                                {
                                    if (times_perturbed[worksheet.Index - 1][row][col] != 0)
                                    {
                                        if (impacts_grid[worksheet.Index - 1][row][col][i] > 2)
                                        {
                                            int[] outlier = new int[3];
                                            outlier[0] = worksheet.Index - 1;
                                            outlier[1] = row;
                                            outlier[2] = col;
                                            outliers.Add(outlier);
                                        }
                                    }
                                }
                            }
                        }
                    }

                    //Find the highest weighted average z-score among the outliers
                    double max_weighted_z_score = 0.0;
                    int[][] outliers_array = outliers.ToArray();
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
                }
                else //if (toggle_weighted_average.Checked)
                {
                    //Find max z-score
                    double max_z_score = 0.0;
                    foreach (Excel.Worksheet worksheet in Globals.ThisAddIn.Application.Worksheets)
                    {
                        for (int row = 0; row < worksheet.UsedRange.Rows.Count; row++)
                        {
                            for (int col = 0; col < worksheet.UsedRange.Columns.Count; col++)
                            {
                                if (average_z_scores[worksheet.Index - 1][row][col] > max_z_score)
                                {
                                    max_z_score = average_z_scores[worksheet.Index - 1][row][col];
                                }
                            }
                        }
                    }

                    //Color based on weighted average z-score being > 2.0
                    foreach (Excel.Worksheet worksheet in Globals.ThisAddIn.Application.Worksheets)
                    {
                        for (int row = 0; row < worksheet.UsedRange.Rows.Count; row++)
                        {
                            for (int col = 0; col < worksheet.UsedRange.Columns.Count; col++)
                            {
                                if (average_z_scores[worksheet.Index - 1][row][col] > 2.0)
                                {
                                    worksheet.Cells[row + 1, col + 1].Interior.Color = System.Drawing.Color.FromArgb(Convert.ToInt32(255 - (average_z_scores[worksheet.Index - 1][row][col] / max_z_score) * 255), 255, 255);
                                }
                            }
                        }
                    }

                    //Excel.Worksheet ws = Globals.ThisAddIn.Application.Worksheets[1];
                    //Excel.Worksheet out_ws = Globals.ThisAddIn.Application.Worksheets[3];
                    //for (int row = 0; row < ws.UsedRange.Rows.Count; row++)
                    //{
                    //    for (int col = 0; col < ws.UsedRange.Columns.Count; col++)
                    //    {
                    //        if (times_perturbed[ws.Index - 1][row][col] != 0)
                    //        {
                    //            for (int i = 0; i < output_cells.Count; i++)
                    //            {
                    //                MessageBox.Show("Worksheet " + ws.Name + ": Cell R" + (row + 1) + "C" + (col + 1) + ": " + impacts_grid[ws.Index - 1][row][col][i]);
                    //                out_ws.Cells[row + 1 * col + 1, 1].Value = "Row: " + (row + 1) + " Col: " + (col + 1);
                    //                out_ws.Cells[(row + 1) * (col + 1), i + 2].Value = impacts_grid[ws.Index - 1][row][col][i];
                    //            }
                    //        }
                    //    }
                    //}
                }
            }

            //if (toggle_global_perturbation.Checked)
            //{
            //    //Divide each influence entry by the number of times perturbed to get a per-swap influence value.
            //    //Also find global max influences by looping over influences_grid
            //    double global_max_inf = 0.0;
            //    foreach (Excel.Worksheet worksheet in Globals.ThisAddIn.Application.Worksheets)
            //    {
            //        for (int row = 0; row < worksheet.UsedRange.Rows.Count; row++)
            //        {
            //            for (int col = 0; col < worksheet.UsedRange.Columns.Count; col++)
            //            {
            //                if (times_perturbed[worksheet.Index - 1][row][col] != 0.0)
            //                {
            //                    influences_grid[worksheet.Index - 1][row][col] = influences_grid[worksheet.Index - 1][row][col] / times_perturbed[worksheet.Index - 1][row][col];
            //                }
            //                if (influences_grid[worksheet.Index - 1][row][col] > global_max_inf)
            //                {
            //                    global_max_inf = influences_grid[worksheet.Index - 1][row][col];
            //                }
            //            }
            //        }
            //    }
            //    //Normalize the influences based on the largest influence values. This way they are >= 0 and < 1.
            //    //Color cells based on influence if global perturbation is checked
            //    foreach (Excel.Worksheet worksheet in Globals.ThisAddIn.Application.Worksheets)
            //    {
            //        for (int row = 0; row < worksheet.UsedRange.Rows.Count; row++)
            //        {
            //            for (int col = 0; col < worksheet.UsedRange.Columns.Count; col++)
            //            {
            //                if (global_max_inf != 0.0)
            //                {
            //                    influences_grid[worksheet.Index - 1][row][col] = influences_grid[worksheet.Index - 1][row][col] / global_max_inf;
            //                }
            //                //Find the cell that is stored in this grid entry
            //                Excel.Range cell = null;
            //                foreach (Excel.Worksheet ws in Globals.ThisAddIn.Application.Worksheets)
            //                {
            //                    if (ws.Index == worksheet.Index)
            //                    {
            //                        cell = ws.Cells[row + 1, col + 1]; //row and column are 1-indexed in ws.Cells
            //                        break;
            //                    }
            //                }
            //                if (!toggle_analyze_outliers.Checked) //if no further analysis to the influence is needed, color the cell
            //                {
            //                    cell.Interior.Color = System.Drawing.Color.FromArgb(Convert.ToInt32(255 - influences_grid[worksheet.Index - 1][row][col] * 255), 255, 255);
            //                }
            //            }
            //        }
            //    }
            //    //if (toggle_analyze_outliers.Checked)
            //    //{
            //    //    Excel.Range range = null;
            //    //    bool first = true;
            //    //    foreach (Excel.Worksheet worksheet in Globals.ThisAddIn.Application.Worksheets)
            //    //    {
            //    //        for (int row = 0; row < worksheet.UsedRange.Rows.Count; row++)
            //    //        {
            //    //            for (int col = 0; col < worksheet.UsedRange.Columns.Count; col++)
            //    //            {
            //    //                if (times_perturbed[worksheet.Index - 1][row][col] != 0.0)
            //    //                {
            //    //                    //Find the cell that is stored in this grid entry
            //    //                    Excel.Range cell = null;
            //    //                    //Finding the right worksheet has to be done this way because a worksheet's index is not the index in the collection Globals.ThisAddIn.Application.Worksheets 
            //    //                    //-- this collection does not include chart sheets, but the index does
            //    //                    foreach (Excel.Worksheet ws in Globals.ThisAddIn.Application.Worksheets)
            //    //                    {
            //    //                        if (ws.Index == worksheet.Index)
            //    //                        {
            //    //                            cell = ws.Cells[row + 1, col + 1]; //row and column are 1-indexed in ws.Cells
            //    //                            if (first)
            //    //                            {
            //    //                                range = cell;
            //    //                                first = false;
            //    //                            }
            //    //                            break;
            //    //                        }
            //    //                    }
            //    //                    range = Globals.ThisAddIn.Application.Union(range, cell, Type.Missing,
            //    //                            Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
            //    //                            Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
            //    //                            Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
            //    //                            Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            //    //                }
            //    //            }
            //    //        }
            //    //    }
            //    //    MessageBox.Show("Running Peirce on range " + range.Address);
            //    //    run_peirce(range);
            //    //    /**
            //    //    //Compute average influence
            //    //    double average_influence = 0.0;
            //    //    double denominator = 0.0;
            //    //    //TODO: if there are overflow issues consider making total_influence an array of doubles (of size 100 for instance) and use each slot as a bin for parts of the sum
            //    //    //each part can be divided by the denominator and then the average_influence is the sum of the entries in the array
            //    //    double total_influence = 0.0;
            //    //    foreach (Excel.Worksheet worksheet in Globals.ThisAddIn.Application.Worksheets)
            //    //    {
            //    //        for (int row = 0; row < worksheet.UsedRange.Rows.Count; row++)
            //    //        {
            //    //            for (int col = 0; col < worksheet.UsedRange.Columns.Count; col++)
            //    //            {
            //    //                if (times_perturbed[worksheet.Index - 1][row][col] != 0.0)
            //    //                {
            //    //                    total_influence += influences_grid[worksheet.Index - 1][row][col];
            //    //                    denominator++;
            //    //                }
            //    //            }
            //    //        }
            //    //    }
            //    //    average_influence = total_influence / denominator;
            //    //    //Compute the standard deviation
            //    //    double variance = 0.0;  //stores the sum of the suqared differences from the mean divided by the denominator
            //    //    foreach (Excel.Worksheet worksheet in Globals.ThisAddIn.Application.Worksheets)
            //    //    {
            //    //        for (int row = 0; row < worksheet.UsedRange.Rows.Count; row++)
            //    //        {
            //    //            for (int col = 0; col < worksheet.UsedRange.Columns.Count; col++)
            //    //            {
            //    //                if (times_perturbed[worksheet.Index - 1][row][col] != 0.0)
            //    //                {
            //    //                    variance += (influences_grid[worksheet.Index - 1][row][col] - average_influence) * (influences_grid[worksheet.Index - 1][row][col] - average_influence) / denominator;
            //    //                }
            //    //            }
            //    //        }
            //    //    }
            //    //    double standard_deviation = Math.Sqrt(variance);                            
            //    //    //Color cells that lie further than two standard deviations away from the mean
            //    //    foreach (Excel.Worksheet worksheet in Globals.ThisAddIn.Application.Worksheets)
            //    //    {
            //    //        for (int row = 0; row < worksheet.UsedRange.Rows.Count; row++)
            //    //        {
            //    //            for (int col = 0; col < worksheet.UsedRange.Columns.Count; col++)
            //    //            {
            //    //                //Find the cell that is stored in this grid entry
            //    //                Excel.Range cell = null;
            //    //                //Finding the right worksheet has to be done this way because a worksheet's index is not the index in the collection Globals.ThisAddIn.Application.Worksheets 
            //    //                //-- this collection does not include chart sheets, but the index does
            //    //                foreach (Excel.Worksheet ws in Globals.ThisAddIn.Application.Worksheets)
            //    //                {
            //    //                    if (ws.Index == worksheet.Index)
            //    //                    {
            //    //                        cell = ws.Cells[row + 1, col + 1]; //row and column are 1-indexed in ws.Cells
            //    //                        break;
            //    //                    }
            //    //                }
            //    //                //If inluence is more than two standard deviations away from the mean, color that cell
            //    //                if (Math.Abs(influences_grid[worksheet.Index - 1][row][col] - average_influence) > 2 * standard_deviation) 
            //    //                {
            //    //                    cell.Interior.Color = System.Drawing.Color.FromArgb(Convert.ToInt32(255 - influences_grid[worksheet.Index - 1][row][col] * 255), 255, 255);
            //    //                }
            //    //                else
            //    //                {
            //    //                    cell.Interior.Color = System.Drawing.Color.White;
            //    //                }
            //    //            }
            //    //        }
            //    //    }
            //    //    **/
            //    //}
            //}
            if (toggle_no_sceen_updating.Checked)
            {
                Globals.ThisAddIn.Application.ScreenUpdating = true;
            }
            sw.Stop();
            // Get the elapsed time as a TimeSpan value.
            //TimeSpan ts = sw.Elapsed;

            // Format and display the TimeSpan value. 
            //string swappingTime = String.Format("{0:00}:{1:00}:{2:00}.{3:00}", ts.Hours, ts.Minutes, ts.Seconds, ts.Milliseconds / 10);
            
            //Display timeDisplay = new Display();
            //timeDisplay.textBox1.Text = "Tree construction time: " + treeTime + "    \nSwapping time: " + swappingTime;
            //timeDisplay.ShowDialog();

            //Procedure for swapping values within ranges, replacing all repeated values at once
            if (checkBox2.Checked) //Checks if the option for swapping values simultaneously is checked
            {
                List<TreeNode> swap_domain;
                if (toggle_array_storage.Checked)
                {
                    swap_domain = ranges;
                }
                else
                {
                    swap_domain = nodes;
                }
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
                            //MessageBox.Show(twin_cells_string);
                            Excel.Range twin_cells = parent.getWorksheetObject().get_Range(twin_cells_string);
                            String[] formulas = new String[twin_count]; //Stores the formulas in the twin_cells
                            int i = 0; //Counter for indexing within the formulas array
                            foreach (Excel.Range cell in twin_cells)
                            {
                                if (cell.HasFormula)
                                {
                                    //MessageBox.Show(cell.Formula);
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
                                    //MessageBox.Show("Substituting " + sibling.getName() 
                                      //+ "\nDelta = |" + starting_outputs[index] + " - " + activeWorksheet.get_Range(n.getName()).Value + "| / " + starting_outputs[index]
                                      //+ " = " + delta
                                      //+ "\nTotal Delta = " + total_delta);
                                }
                            }
                            total_delta = total_delta / (node.getParents().Count - 1);
                            total_delta = total_delta / twin_count;
                            influences[influence_index] = total_delta;
                            influence_index++;
                            //MessageBox.Show(twin_cells.get_Address() + " Total delta = " + (total_delta * 100) + "%");
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
                        //MessageBox.Show("MIN DELTA: " + min_total_delta + "\nMAX DELTA: " + max_total_delta);
                        foreach (TreeNode parent in node.getParents())
                        {
                            if (max_total_delta != 0)
                            {
                                influences[ind] = (influences[ind] - min_total_delta) / max_total_delta;
                                //MessageBox.Show("Influence = " + influences[ind]);
                            }
                            ind++;
                        }
                        //for (int i = 0; i < node.getParents().Count; i++)
                        //{
                        //    if (max_total_delta != 0)
                        //    {
                        //        influences[i] = (influences[i] - min_total_delta) / max_total_delta;
                        //    }
                        //}
                        int indexer = 0;
                        foreach (TreeNode parent in node.getParents())
                        {
                            Excel.Range cell = parent.getWorksheetObject().get_Range(parent.getName());
                            //parent.setOriginalColor(System.Drawing.ColorTranslator.FromOle((int)cell.Interior.Color));
                            //Only store the original color on the first run of the tool
                            //if (toolHasNotRun == false)
                            //{
                                //parent.setOriginalColor(cell.Interior.ColorIndex);
                            //    parent.setOriginalColor(System.Drawing.ColorTranslator.FromOle((int)cell.Interior.Color));
                            //}
                            cell.Interior.Color = System.Drawing.Color.FromArgb(Convert.ToInt32(255 - influences[indexer] * 255), 255, 255);
                            indexer++;
                        }
                    }
                }
            }


            ////Print out text for GraphViz representation of the dependence graph
            //string tree = "";
            //string ranges_text = "";
            //if (toggle_array_storage.Checked)
            //{
            //    foreach (TreeNode[][] node_arr_arr in nodes_grid)
            //    {
            //        foreach (TreeNode[] node_arr in node_arr_arr)
            //        {
            //            foreach (TreeNode node in node_arr)
            //            {
            //                if (node != null)
            //                {
            //                    tree += node.toGVString(0) + "\n"; //tree += node.toGVString(max_weight) + "\n";
            //                }
            //            }
            //        }
            //    }
            //    foreach (TreeNode node in ranges)
            //    {
            //        tree += node.toGVString(0) + "\n"; //tree += node.toGVString(max_weight) + "\n";
            //        foreach (TreeNode parent in node.getParents())
            //        {
            //            ranges_text += parent.getWorksheetObject().Index + "," + parent.getName().Replace("$","") + "," + parent.getWorksheetObject().get_Range(parent.getName()).Value +"\n";
            //        }
            //    }
            //}
            //else //toggle_array_storage not checked
            //{
            //    foreach (TreeNode node in nodes)
            //    {
            //        tree += node.toGVString(0) + "\n"; //tree += node.toGVString(max_weight) + "\n";
            //    }
            //}
            //Display disp = new Display();
            //disp.textBox1.Text = "digraph g{" + tree + "}";
            //disp.ShowDialog();
            //Display disp_ranges = new Display();
            //disp_ranges.textBox1.Text = ranges_text;
            //disp_ranges.ShowDialog();
        }

        List<TreeNode> originalColorNodes;
        
        //Action for "Analyze Worksheet" button
        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            //ProgressBar pb = new ProgressBar();
            //pb.Style = ProgressBarStyle.Marquee;
            //pb.MarqueeAnimationSpeed = 100;
            //pb.Visible = true; 
            //pb.Show();
            //IdentifyRanges();

            //Construct a new tree every time the tool is run
            nodes = new List<TreeNode>();        //This is a list holding all the TreeNodes in the Excel file

            if (toggle_array_storage.Checked)
            {
                ranges = new List<TreeNode>();        //This is a list holding all the ranges of TreeNodes in the Excel file
                charts = new List<TreeNode>();        //This is a list holding all the chart TreeNodes in the Excel file
                nodes_grid = new TreeNode[Globals.ThisAddIn.Application.Worksheets.Count + Globals.ThisAddIn.Application.Charts.Count][][];
                int index = 0;
                foreach (Excel.Worksheet worksheet in Globals.ThisAddIn.Application.Worksheets)
                {
                    nodes_grid[worksheet.Index - 1] = new TreeNode[worksheet.UsedRange.Rows.Count][];
                    for (int row = 0; row < worksheet.UsedRange.Rows.Count; row++)
                    {
                        nodes_grid[worksheet.Index - 1][row] = new TreeNode[worksheet.UsedRange.Columns.Count];
                        for (int col = 0; col < worksheet.UsedRange.Columns.Count; col++)
                        {
                            nodes_grid[worksheet.Index - 1][row][col] = null;
                        }
                    }
                    index++;
                }
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

            if (toolHasNotRun)
            {
                //Save starting colors 
                originalColorNodes = new List<TreeNode>(); 
                foreach (Excel.Worksheet worksheet in Globals.ThisAddIn.Application.Worksheets)
                {
                    foreach (Excel.Range cell in worksheet.UsedRange)
                    {
                        TreeNode n = new TreeNode(cell.Address, cell.Worksheet.Name);  //Create a TreeNode for every cell with the name being the cell's address and set the node's worksheet appropriately
                        //n.setOriginalColor(cell.Interior.ColorIndex);
                        n.setOriginalColor(System.Drawing.ColorTranslator.FromOle((int)cell.Interior.Color));
                        originalColorNodes.Add(n);
                    }
                }
                constructTree();
                toolHasNotRun = false;
            }
            else
            {
                constructTree();
            }
        }

        private void dropDown1_SelectionChanged(object sender, RibbonControlEventArgs e)
        {
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

        private void button7_Click(object sender, RibbonControlEventArgs e)
        {
            foreach (Excel.Name n in Globals.ThisAddIn.Application.Names)
            {
                MessageBox.Show(n.Name + " : " + n.RefersToRange.Address);
            }
            //String[] worksheet_names = new String[Globals.ThisAddIn.Application.Worksheets.Count];
            //MessageBox.Show("A1 row: " + Globals.ThisAddIn.Application.Worksheets[1].get_Range("A1").Row);
            //MessageBox.Show("A1 col: " + Globals.ThisAddIn.Application.Worksheets[1].get_Range("A1").Column);
            nodes_grid = new TreeNode[Globals.ThisAddIn.Application.Worksheets.Count][][];
            int index = 0;
            foreach (Excel.Worksheet worksheet in Globals.ThisAddIn.Application.Worksheets)
            {
                MessageBox.Show("Worksheet " + worksheet.Name + " index: " + worksheet.Index);
                nodes_grid[index] = new TreeNode[worksheet.UsedRange.Rows.Count][];
                for (int row = 0; row < worksheet.UsedRange.Rows.Count; row++)
                {
                    nodes_grid[index][row] = new TreeNode[worksheet.UsedRange.Columns.Count];
                    for (int col = 0; col < worksheet.UsedRange.Rows.Count; col++)
                    {
                        nodes_grid[index][row][col] = null;
                    }
                }
                index++;
            }
            Excel.Worksheet activeWorksheet = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet;
            /**
            foreach (string s in worksheet_names)
            {
                foreach (Excel.Range c in activeWorksheet.UsedRange)
                {
                    string formula = c.Formula;  //The formula contained in the cell
                    string worksheet_name = s.Replace("+", @"\+").Replace("^", @"\^").Replace("$", @"\$").Replace(".", @"\.");
                    //First look for range references in the formula
                    MatchCollection matchedRanges = Regex.Matches(formula, @"('" + worksheet_name + @"'!\$?[A-Z]+\$?[1-9]\d*:\$?[A-Z]+\$?[1-9]\d*)");  //A collection of all the range references in the formula; each item is a range reference such as A1:A10
                    foreach (Match match in matchedRanges)
                    {
                        formula = formula.Replace(match.Value, ""); //remove any identified matches so that they are not counted again later
                        //MessageBox.Show(match.Value);
                    }
                    matchedRanges = Regex.Matches(formula, @"(" + worksheet_name + @"!\$?[A-Z]+\$?[1-9]\d*:\$?[A-Z]+\$?[1-9]\d*)");  //A collection of all the range references in the formula; each item is a range reference such as A1:A10
                    foreach (Match match in matchedRanges)
                    {
                        formula = formula.Replace(match.Value, ""); //remove any identified matches so that they are not counted again later
                        //MessageBox.Show(match.Value);
                    }

                    matchedRanges = Regex.Matches(formula, @"('" + worksheet_name + @"'!\$?[A-Z]+\$?[1-9]\d*)");
                    foreach (Match match in matchedRanges)
                    {
                        formula = formula.Replace(match.Value, ""); //remove any identified matches so that they are not counted again later
                        //MessageBox.Show(match.Value);
                    }
                    matchedRanges = Regex.Matches(formula, @"(" + worksheet_name + @"!\$?[A-Z]+\$?[1-9]\d*)");
                    foreach (Match match in matchedRanges)
                    {
                        formula = formula.Replace(match.Value, ""); //remove any identified matches so that they are not counted again later
                        //MessageBox.Show(match.Value);
                    }
                }
            }
             **/
        }

        //Action for "Clear coloring" button
        private void button8_Click(object sender, RibbonControlEventArgs e)
        {
            foreach (TreeNode node in originalColorNodes)
            {
                //If the node is just a cell, clear any coloring
                if (!node.isChart() && !node.isRange())
                {
                    //MessageBox.Show(node.getName() + " " + node.getOriginalColor());
                    //node.getWorksheetObject().get_Range(node.getName()).Interior.ColorIndex = 0;
                    //node.getWorksheetObject().get_Range(node.getName()).Interior.ColorIndex = node.getOriginalColor();
                    node.getWorksheetObject().get_Range(node.getName()).Interior.Color = node.getOriginalColor();
                }
            }
            //After having cleared things we would like to be able to run the tool again. 
            toolHasNotRun = true;
        }

        private void toggle_speed_Click(object sender, RibbonControlEventArgs e)
        {

        }

        private void peirce_button_Click(object sender, RibbonControlEventArgs e)
        {
            //run_peirce(Globals.ThisAddIn.Application.Selection as Excel.Range);
            //get_peirce_cutoff((Globals.ThisAddIn.Application.Selection as Excel.Range).Cells.Count);
            //HERE HERE
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
                            //MessageBox.Show(cell.Address + " is an outlier.");
                            cell.Interior.Color = System.Drawing.Color.Red;
                            outliers.Add(cell);
                            count_rejected++;
                        }
                    }
                }
                k = k + count_rejected;
                //MessageBox.Show("Got here.");
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
                            //MessageBox.Show(cell.Address + " is an outlier.");
                            outliers.Add(d);
                            count_rejected++;
                        }
                    }
                }
                k = k + count_rejected;
                //MessageBox.Show("Got here.");
            } while (count_rejected > 0);
            return outliers;
        }

        private void find_references(TreeNode formula_node, string formula)
        {
            MatchCollection matchedRanges = null;
            MatchCollection matchedCells = null;
            int ws_index = 1;
            foreach (Excel.Worksheet worksheet in Globals.ThisAddIn.Application.Worksheets)
            {
                string s = worksheet.Name;
                string worksheet_name = s.Replace("+", @"\+").Replace("^", @"\^").Replace("$", @"\$").Replace(".", @"\."); //Escape certain characters in the regular expression
                //First look for range references of the form 'worksheet_name'!A1:A10 in the formula (with quotation marks around the name)
                if (toggle_compile_regex.Checked)
                {
                    matchedRanges = regex_array[4 * (ws_index - 1)].Matches(formula);
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
                    if (toggle_array_storage.Checked)
                    {
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
                    }
                    else //toggle_array_storage not checked
                    {
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
                    }
                    //If it was not found, create it
                    if (range == null)
                    {
                        range = new TreeNode(endCells[0].Replace("$", "") + "_to_" + endCells[1].Replace("$", ""), ws_name);
                        //MessageBox.Show("Created range node:" + ws_name + "_" + endCells[0] + ":" + endCells[1]);
                        if (toggle_array_storage.Checked)
                        {
                            ranges.Add(range);
                        }
                        else //toggle_array_storage not checked
                        {
                            nodes.Add(range);
                        }
                    }
                    formula_node.addParent(range);
                    range.addChild(formula_node);
                    //Add each cell contained in the range to the dependencies
                    foreach (Excel.Range cellInRange in Globals.ThisAddIn.Application.Worksheets[ws_index].Range[endCells[0], endCells[1]])
                    {
                        TreeNode input_cell = null;
                        //Find the node object for the current cell in the existing TreeNodes
                        if (toggle_array_storage.Checked)
                        {
                            //Check if this cell's coordinates are within the bounds of the used range, otherwise there will be an index out of bounds error
                            if (cellInRange.Column <= cellInRange.Worksheet.UsedRange.Columns.Count && cellInRange.Row <= cellInRange.Worksheet.UsedRange.Rows.Count)
                            {
                                //if a TreeNode exists for this cell already
                                if (nodes_grid[cellInRange.Worksheet.Index - 1][cellInRange.Row - 1][cellInRange.Column - 1] != null)
                                {
                                    input_cell = nodes_grid[cellInRange.Worksheet.Index - 1][cellInRange.Row - 1][cellInRange.Column - 1];
                                }
                            }
                        }
                        else  //toggle_array_storage not checked
                        {
                            foreach (TreeNode node in nodes)
                            {
                                if (node.getName().Replace("$", "") == cellInRange.Address.Replace("$", "") && node.getWorksheet() == cellInRange.Worksheet.Name)
                                {
                                    input_cell = node;
                                }
                                else
                                    continue;
                            }
                        }
                        //If it wasn't found, then it is blank, and we have to create a TreeNode for it
                        if (input_cell == null)
                        {
                            input_cell = new TreeNode(cellInRange.Address, cellInRange.Worksheet.Name);
                            if (toggle_array_storage.Checked)
                            {
                                //Check if this cell's coordinates are within the bounds of the used range, otherwise there will be an index out of bounds error
                                if (cellInRange.Column <= cellInRange.Worksheet.UsedRange.Columns.Count && cellInRange.Row <= cellInRange.Worksheet.UsedRange.Rows.Count)
                                {
                                    nodes_grid[cellInRange.Worksheet.Index - 1][cellInRange.Row - 1][cellInRange.Column - 1] = input_cell;
                                }
                            }
                            else  //toggle_array_storage not checked
                            {
                                nodes.Add(input_cell);
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
                    matchedRanges = regex_array[4 * (ws_index - 1) + 1].Matches(formula);
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
                    if (toggle_array_storage.Checked)
                    {
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
                    }
                    else //toggle_array_storage not checked
                    {
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
                    }
                    //If it does not exist, create it
                    if (range == null)
                    {
                        range = new TreeNode(endCells[0].Replace("$", "") + "_to_" + endCells[1].Replace("$", ""), ws_name);
                        //MessageBox.Show("Created node:" + ws_name + "_" + endCells[0] + ":" + endCells[1]);
                        if (toggle_array_storage.Checked)
                        {
                            ranges.Add(range);
                        }
                        else //toggle_array_storage not checked
                        {
                            nodes.Add(range);
                        }
                    }
                    formula_node.addParent(range);
                    range.addChild(formula_node);
                    //Add each cell contained in the range to the dependencies
                    foreach (Excel.Range cellInRange in Globals.ThisAddIn.Application.Worksheets[ws_index].Range[endCells[0], endCells[1]])
                    {
                        TreeNode input_cell = null;
                        //Find the node object for the current cell in the existing TreeNodes
                        if (toggle_array_storage.Checked)
                        {
                            //Check if this cell's coordinates are within the bounds of the used range, otherwise there will be an index out of bounds error
                            if (cellInRange.Column <= cellInRange.Worksheet.UsedRange.Columns.Count && cellInRange.Row <= cellInRange.Worksheet.UsedRange.Rows.Count)
                            {
                                //if a TreeNode exists for this cell already
                                if (nodes_grid[cellInRange.Worksheet.Index - 1][cellInRange.Row - 1][cellInRange.Column - 1] != null)
                                {
                                    input_cell = nodes_grid[cellInRange.Worksheet.Index - 1][cellInRange.Row - 1][cellInRange.Column - 1];
                                }
                            }
                        }
                        else  //toggle_array_storage not checked
                        {
                            foreach (TreeNode node in nodes)
                            {
                                if (node.getName().Replace("$", "") == cellInRange.Address.Replace("$", "") && node.getWorksheet() == cellInRange.Worksheet.Name)
                                {
                                    input_cell = node;
                                }
                                else
                                    continue;
                            }
                        }
                        //If it wasn't found, then it is blank, and we have to create a TreeNode for it
                        if (input_cell == null)
                        {
                            input_cell = new TreeNode(cellInRange.Address, cellInRange.Worksheet.Name);
                            if (toggle_array_storage.Checked)
                            {
                                //Check if this cell's coordinates are within the bounds of the used range, otherwise there will be an index out of bounds error
                                if (cellInRange.Column <= cellInRange.Worksheet.UsedRange.Columns.Count && cellInRange.Row <= cellInRange.Worksheet.UsedRange.Rows.Count)
                                {
                                    nodes_grid[cellInRange.Worksheet.Index - 1][cellInRange.Row - 1][cellInRange.Column - 1] = input_cell;
                                }
                            }
                            else  //toggle_array_storage not checked
                            {
                                nodes.Add(input_cell);
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
                    matchedCells = regex_array[4 * (ws_index - 1) + 2].Matches(formula);
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
                    if (toggle_array_storage.Checked)
                    {
                        //Check if this cell's coordinates are within the bounds of the used range of its spreadsheet, otherwise there will be an index out of bounds error
                        if (input.Column <= input.Worksheet.UsedRange.Columns.Count && input.Row <= input.Worksheet.UsedRange.Rows.Count)
                        {
                            //if a TreeNode exists for this cell already, use it
                            if (nodes_grid[input.Worksheet.Index - 1][input.Row - 1][input.Column - 1] != null)
                            {
                                input_cell = nodes_grid[input.Worksheet.Index - 1][input.Row - 1][input.Column - 1];
                            }
                        }
                    }
                    else //toggle_array_storage not checked
                    {
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
                    }
                    //If it wasn't found, then it is blank, and we have to create a TreeNode for it
                    if (input_cell == null)
                    {
                        input_cell = new TreeNode(cell_coordinates.Replace("$", ""), ws_name);
                        if (toggle_array_storage.Checked)
                        {
                            //Check if this cell's coordinates are within the bounds of the used range, otherwise there will be an index out of bounds error
                            if (input.Column <= input.Worksheet.UsedRange.Columns.Count && input.Row <= input.Worksheet.UsedRange.Rows.Count)
                            {
                                nodes_grid[input.Worksheet.Index - 1][input.Row - 1][input.Column - 1] = input_cell;
                            }
                        }
                        else //toggle_array_storage not checked
                        {
                            nodes.Add(input_cell);
                        }
                    }

                    //Update the dependencies
                    formula_node.addParent(input_cell);
                    input_cell.addChild(formula_node);
                }

                //Lastly we look for references of the kind worksheet_name!A1 (without quotation marks)
                if (toggle_compile_regex.Checked)
                {
                    matchedCells = regex_array[4 * (ws_index - 1) + 3].Matches(formula);
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
                    if (toggle_array_storage.Checked)
                    {
                        //Check if this cell's coordinates are within the bounds of the used range of its spreadsheet, otherwise there will be an index out of bounds error
                        if (input.Column <= input.Worksheet.UsedRange.Columns.Count && input.Row <= input.Worksheet.UsedRange.Rows.Count)
                        {
                            //if a TreeNode exists for this cell already, use it
                            if (nodes_grid[input.Worksheet.Index - 1][input.Row - 1][input.Column - 1] != null)
                            {
                                input_cell = nodes_grid[input.Worksheet.Index - 1][input.Row - 1][input.Column - 1];
                            }
                        }
                    }
                    else //toggle_array_storage not checked
                    {
                        foreach (TreeNode node in nodes)
                        {
                            if (node.getName().Replace("$", "") == cell_coordinates.Replace("$", "") && node.getWorksheet().Equals(ws_name))
                            {
                                input_cell = node;
                            }
                            else
                            {
                                continue;
                            }
                        }
                    }
                    //If it wasn't found, then it is blank, and we have to create a TreeNode for it
                    if (input_cell == null)
                    {
                        input_cell = new TreeNode(cell_coordinates.Replace("$", ""), ws_name);
                        if (toggle_array_storage.Checked)
                        {
                            //Check if this cell's coordinates are within the bounds of the used range, otherwise there will be an index out of bounds error
                            if (input.Column <= input.Worksheet.UsedRange.Columns.Count && input.Row <= input.Worksheet.UsedRange.Rows.Count)
                            {
                                nodes_grid[input.Worksheet.Index - 1][input.Row - 1][input.Column - 1] = input_cell;
                            }
                        }
                        else //toggle_array_storage not checked
                        {
                            nodes.Add(input_cell);
                        }
                    }

                    //Update the dependencies
                    formula_node.addParent(input_cell);
                    input_cell.addChild(formula_node);
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
            foreach (Match match in matchedRanges)
            {
                formula = formula.Replace(match.Value, "");
                string[] endCells = match.Value.Split(':');     //Split up each matched range into the start and end cells of the range
                TreeNode range = null;
                //Try to find the range in existing TreeNodes
                if (toggle_array_storage.Checked)
                {
                    foreach (TreeNode n in ranges)
                    {
                        if (n.getName().Replace("$", "") == endCells[0].Replace("$", "") + "_to_" + endCells[1].Replace("$", "") && n.getWorksheet() == formula_node.getWorksheetObject().Name)
                        {
                            range = n;
                        }
                        else
                        {
                            continue;
                        }
                    }
                }
                else //toggle_array_storage not checked
                {
                    foreach (TreeNode n in nodes)
                    {
                        if (n.getName().Replace("$", "") == endCells[0].Replace("$", "") + "_to_" + endCells[1].Replace("$", "") && n.getWorksheet() == formula_node.getWorksheetObject().Name)
                        {
                            range = n;
                        }
                        else
                        {
                            continue;
                        }
                    }
                }
                //If it does not exist, create it
                if (range == null)
                {
                    //MessageBox.Show("Created range node:" + c.Worksheet.Name + "_" + endCells[0] + ":" + endCells[1]);
                    range = new TreeNode(endCells[0].Replace("$", "") + "_to_" + endCells[1].Replace("$", ""), formula_node.getWorksheetObject().Name);
                    if (toggle_array_storage.Checked)
                    {
                        ranges.Add(range);
                    }
                    else
                    {
                        nodes.Add(range);
                    }
                }
                formula_node.addParent(range);
                range.addChild(formula_node);
                //Add each cell contained in the range to the dependencies
                foreach (Excel.Range cellInRange in formula_node.getWorksheetObject().Range[endCells[0], endCells[1]])
                {
                    TreeNode input_cell = null;
                    //Find the node object for the current cell in the existing TreeNodes
                    //HERE HERE
                    if (toggle_array_storage.Checked)
                    {
                        //Check if this cell's coordinates are within the bounds of the used range, otherwise there will be an index out of bounds error
                        if (cellInRange.Column <= cellInRange.Worksheet.UsedRange.Columns.Count && cellInRange.Row <= cellInRange.Worksheet.UsedRange.Rows.Count)
                        {
                            //if a TreeNode exists for this cell already, use it
                            if (nodes_grid[cellInRange.Worksheet.Index - 1][cellInRange.Row - 1][cellInRange.Column - 1] != null)
                            {
                                input_cell = nodes_grid[cellInRange.Worksheet.Index - 1][cellInRange.Row - 1][cellInRange.Column - 1];
                            }
                        }
                    }
                    else //toggle_array_storage not checked
                    {
                        foreach (TreeNode node in nodes)
                        {
                            if (node.getName().Replace("$", "") == cellInRange.Address.Replace("$", "") && node.getWorksheet() == cellInRange.Worksheet.Name)
                            {
                                input_cell = node;
                            }
                            else
                            {
                                continue;
                            }
                        }
                    }
                    //If it wasn't found, then it is blank, and we have to create a TreeNode for it
                    if (input_cell == null)
                    {
                        input_cell = new TreeNode(cellInRange.Address, cellInRange.Worksheet.Name);
                        if (toggle_array_storage.Checked)
                        {
                            //Check if this cell's coordinates are within the bounds of the used range, otherwise there will be an index out of bounds error
                            if (cellInRange.Column <= cellInRange.Worksheet.UsedRange.Columns.Count && cellInRange.Row <= cellInRange.Worksheet.UsedRange.Rows.Count)
                            {
                                nodes_grid[cellInRange.Worksheet.Index - 1][cellInRange.Row - 1][cellInRange.Column - 1] = input_cell;
                            }
                        }
                        else //toggle_array_storage not checked
                        {
                            nodes.Add(input_cell);
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

                string[] endCells = named_range.RefersToRange.Address.Split(':');     //Split up each named range into the start and end cells of the range
                TreeNode range = null;
                //Try to find the range in existing TreeNodes
                if (toggle_array_storage.Checked)
                {
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
                }
                else //toggle_array_storage not checked
                {
                    foreach (TreeNode n in nodes)
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
                }
                //If it does not exist, create it
                if (range == null)
                {
                    //MessageBox.Show("Created range node:" + c.Worksheet.Name + "_" + endCells[0] + ":" + endCells[1]);
                    range = new TreeNode(endCells[0].Replace("$", "") + "_to_" + endCells[1].Replace("$", ""), named_range.RefersToRange.Worksheet.Name);
                    if (toggle_array_storage.Checked)
                    {
                        ranges.Add(range);
                    }
                    else
                    {
                        nodes.Add(range);
                    }
                }
                formula_node.addParent(range);
                range.addChild(formula_node);
                //Add each cell contained in the range to the dependencies
                foreach (Excel.Range cellInRange in named_range.RefersToRange.Worksheet.Range[endCells[0], endCells[1]])
                {
                    TreeNode input_cell = null;
                    //Find the node object for the current cell in the existing TreeNodes
                    if (toggle_array_storage.Checked)
                    {
                        //Check if this cell's coordinates are within the bounds of the used range, otherwise there will be an index out of bounds error
                        if (cellInRange.Column <= cellInRange.Worksheet.UsedRange.Columns.Count && cellInRange.Row <= cellInRange.Worksheet.UsedRange.Rows.Count)
                        {
                            //if a TreeNode exists for this cell already, use it
                            if (nodes_grid[cellInRange.Worksheet.Index - 1][cellInRange.Row - 1][cellInRange.Column - 1] != null)
                            {
                                input_cell = nodes_grid[cellInRange.Worksheet.Index - 1][cellInRange.Row - 1][cellInRange.Column - 1];
                            }
                        }
                    }
                    else //toggle_array_storage not checked
                    {
                        foreach (TreeNode node in nodes)
                        {
                            if (node.getName().Replace("$", "") == cellInRange.Address.Replace("$", "") && node.getWorksheet() == cellInRange.Worksheet.Name)
                            {
                                input_cell = node;
                            }
                            else
                                continue;
                        }
                    }
                    //If it wasn't found, then it is blank, and we have to create a TreeNode for it
                    if (input_cell == null)
                    {
                        input_cell = new TreeNode(cellInRange.Address, cellInRange.Worksheet.Name);
                        if (toggle_array_storage.Checked)
                        {
                            //Check if this cell's coordinates are within the bounds of the used range, otherwise there will be an index out of bounds error
                            if (cellInRange.Column <= cellInRange.Worksheet.UsedRange.Columns.Count && cellInRange.Row <= cellInRange.Worksheet.UsedRange.Rows.Count)
                            {
                                nodes_grid[cellInRange.Worksheet.Index - 1][cellInRange.Row - 1][cellInRange.Column - 1] = input_cell;
                            }
                        }
                        else //toggle_array_storage not checked
                        {
                            nodes.Add(input_cell);
                        }
                    }

                    //Update the dependencies
                    range.addParent(input_cell);
                    input_cell.addChild(range);
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
                Excel.Range input = formula_node.getWorksheetObject().get_Range(m.Value);
                TreeNode input_cell = null;
                //MessageBox.Show(m.Value);
                //Find the node object for the current cell in the existing TreeNodes
                if (toggle_array_storage.Checked)
                {
                    //Check if this cell's coordinates are within the bounds of the used range, otherwise there will be an index out of bounds error
                    if (input.Column <= input.Worksheet.UsedRange.Columns.Count && input.Row <= input.Worksheet.UsedRange.Rows.Count)
                    {
                        //if a TreeNode exists for this cell already, use it
                        if (nodes_grid[input.Worksheet.Index - 1][input.Row - 1][input.Column - 1] != null)
                        {
                            input_cell = nodes_grid[input.Worksheet.Index - 1][input.Row - 1][input.Column - 1];
                        }
                    }
                }
                else //toggle_array_storage not checked
                {
                    foreach (TreeNode node in nodes)
                    {
                        if (node.getName().Replace("$", "") == m.Value.Replace("$", "") && node.getWorksheet() == formula_node.getWorksheetObject().Name)
                        {
                            input_cell = node;
                        }
                        else
                        {
                            continue;
                        }
                    }
                }
                //If it wasn't found, then it is blank, and we have to create a TreeNode for it
                if (input_cell == null)
                {
                    input_cell = new TreeNode(m.Value.Replace("$", ""), formula_node.getWorksheetObject().Name);
                    if (toggle_array_storage.Checked)
                    {
                        //Check if this cell's coordinates are within the bounds of the used range, otherwise there will be an index out of bounds error
                        if (input.Column <= input.Worksheet.UsedRange.Columns.Count && input.Row <= input.Worksheet.UsedRange.Rows.Count)
                        {
                            nodes_grid[input.Worksheet.Index - 1][input.Row - 1][input.Column - 1] = input_cell;
                        }
                    }
                    else //toggle_array_storage not checked
                    {
                        nodes.Add(input_cell);
                    }
                }
                //Update the dependencies
                formula_node.addParent(input_cell);
                input_cell.addChild(formula_node);
            }
        }
    }
}
