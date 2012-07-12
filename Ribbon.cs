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
                specialCellFormulas.Interior.Color = System.Drawing.Color.Crimson;
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

        /*
         * This method constructs the dependency graph from the worksheet.
         * It analyzes formulas and looks for references to cells or ranges of cells.
         * It also looks for any charts, and adds those to the dependency graph as well. 
         * After the dependency graph is constructed, we use it to determine and propagate weights to all nodes in the graph. 
         * In the end, a text representation of the dependency graph is given in GraphViz format. It includes the entire graph and the weights of the nodes.
         */
        private void constructTree()
        {
            Excel.Worksheet activeWorksheet = ((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet);
            List<TreeNode> nodes = new List<TreeNode>();        //This is a list holding all the TreeNodes
            Excel.Range analysisRange; //This keeps track of the range to be analyzed - it is either the user's selection or the whole worksheet
            if (checkBox1.Checked)
            {
                analysisRange = Globals.ThisAddIn.Application.Selection as Excel.Range;
            }
            else
            {
                analysisRange = activeWorksheet.UsedRange;
            }
            //First we create nodes for every non-null cell; then we will operate on these node objects, connecting them in the tree, etc. 
            //This includes cells that contain constants and formulas
            foreach (Excel.Range cell in analysisRange)
            {
                if (cell.Value != null)
                {
                    TreeNode n = new TreeNode(cell.Address);  //Create a TreeNode for every cell with the name being the cell's address. 
                    nodes.Add(n);
                }
            }

            //Next we go through the cells that contain formulas in order to extract the dependencies between them and their inputs
            //For every cell that contains a formula, we get the node we created for that cell. Then we parse the formula using a regular expresion 
            //to find any references to cells or ranges. (We first look for references to ranges, because they supersede the single cell references.)
            //Whenever a reference is found, we update the parent-child relationship between the formula cell and the referenced cell or range.
            //If a range reference is found, we create a node representing that range, and we also create nodes for all of the cells that compose it. 
            //The range is connected to the formula cell, and the composing cells are connected to the range. 
            //If a single cell reference is found, we connect it to the formula cell directly. 
            foreach (Excel.Range c in analysisRange)
            {
                if (c.HasFormula)
                {
                    TreeNode formula_cell = null;
                    //Look for the node object for the current cell in the list of TreeNodes
                    foreach (TreeNode n in nodes)
                    {
                        if (n.getName() == c.Address)
                        {
                            formula_cell = n;
                        }
                        else
                        {
                            continue;
                        }
                    }

                    string patternRange = @"(\$?[A-Z]+\$?[1-9]\d*:\$?[A-Z]+\$?[1-9]\d*)";  //Regex for matching range references in formulas such as A1:A10, or $A$1:$A$10 etc.
                    string patternCell = @"(\$?[A-Z]+\$?[1-9]\d*)";        //Regex for matching single cell references such as A1 or $A$1, etc. 
                    string formula = c.Formula;  //The formula contained in the cell

                    //First look for range references in the formula
                    MatchCollection matchedRanges = Regex.Matches(formula, patternRange);  //A collection of all the range references in the formula; each item is a range reference such as A1:A10
                    List<Excel.Range> rangeList = new List<Excel.Range>();
                    foreach (Match match in matchedRanges)
                    {
                        formula = formula.Replace(match.Value, "");
                        string[] endCells = match.Value.Split(':');     //Split up each matched range into the start and end cells of the range
                        TreeNode range = null;
                        //Try to find the range in existing TreeNodes
                        foreach (TreeNode n in nodes)
                        {
                            if (n.getName() == endCells[0].Replace("$", "") + "_to_" + endCells[1].Replace("$", ""))
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
                            range = new TreeNode(endCells[0].Replace("$", "") + "_to_" + endCells[1].Replace("$", ""));
                            nodes.Add(range);
                        }
                        formula_cell.addParent(range);
                        range.addChild(formula_cell);
                        //Add each cell contained in the range to the dependencies
                        foreach (Excel.Range cellInRange in activeWorksheet.Range[endCells[0], endCells[1]])
                        {
                            TreeNode input_cell = null;
                            //Find the node object for the current cell in the list of TreeNodes
                            foreach (TreeNode node in nodes)
                            {
                                if (node.getName() == cellInRange.Address)
                                {
                                    input_cell = node;
                                }
                                else
                                    continue;
                            }
                            //If it wasn't found, then it is blank, and we have to create a TreeNode for it
                            if (input_cell == null)
                            {
                                input_cell = new TreeNode(cellInRange.Address);
                                nodes.Add(input_cell);
                            }

                            //Update the dependencies
                            range.addParent(input_cell);
                            input_cell.addChild(range);
                        }
                    }

                    MatchCollection matchedCells = Regex.Matches(formula, patternCell);  //matchedCells is a collection of all the cells that are referenced by the formula
                    foreach (Match m in matchedCells)
                    {
                        TreeNode input_cell = null;
                        //Find the node object for the current cell in the list of TreeNodes
                        foreach (TreeNode node in nodes)
                        {
                            if (node.getName().Replace("$", "") == m.Value.Replace("$", ""))
                            {
                                input_cell = node;
                            }
                            else
                            {
                                continue;
                            }
                        }

                        //Update the dependencies
                        formula_cell.addParent(input_cell);
                        input_cell.addChild(formula_cell);
                    }
                }
            }

            foreach (Excel.ChartObject chart in (Excel.ChartObjects)activeWorksheet.ChartObjects(Type.Missing))
            {
                TreeNode n = new TreeNode("Chart" + chart.Name.Replace(" ", ""));
                nodes.Add(n);
                foreach (Excel.Series series in (Excel.SeriesCollection)chart.Chart.SeriesCollection(Type.Missing))
                {
                    //MessageBox.Show(series.Formula);
                    string patternRange = @"(\$?[A-Z]+\$?[1-9]\d*:\$?[A-Z]+\$?[1-9]\d*)";  //Regex for matching range references in formulas such as A1:A10, or $A$1:$A$10 etc.
                    string patternCell = @"(\$?[A-Z]+\$?[1-9]\d*)";        //Regex for matching single cell references such as A1 or $A$1, etc. 
                    string formula = series.Formula;  //The formula contained in the cell

                    //First look for range references in the formula
                    MatchCollection matchedRanges = Regex.Matches(formula, patternRange);  //A collection of all the range references in the formula; each item is a range reference such as A1:A10
                    List<Excel.Range> rangeList = new List<Excel.Range>();
                    foreach (Match match in matchedRanges)
                    {
                        formula = formula.Replace(match.Value, "");
                        string[] endCells = match.Value.Split(':');     //Split up each matched range into the start and end cells of the range
                        TreeNode range = null;
                        //Try to find the range in existing TreeNodes
                        foreach (TreeNode node in nodes)
                        {
                            if (node.getName() == endCells[0].Replace("$", "") + "_to_" + endCells[1].Replace("$", ""))
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
                            range = new TreeNode(endCells[0].Replace("$", "") + "_to_" + endCells[1].Replace("$", ""));
                            nodes.Add(range);
                        }
                        n.addParent(range);
                        range.addChild(n);
                        //Add each cell contained in the range to the dependencies
                        foreach (Excel.Range cellInRange in activeWorksheet.Range[endCells[0], endCells[1]])
                        {
                            TreeNode input_cell = null;
                            //Find the node object for the current cell in the list of TreeNodes
                            foreach (TreeNode node in nodes)
                            {
                                if (node.getName().Replace("$", "") == cellInRange.Address.Replace("$", ""))
                                {
                                    input_cell = node;
                                }
                                else
                                    continue;
                            }
                            //If it wasn't found, then it is blank, and we have to create a TreeNode for it
                            if (input_cell == null)
                            {
                                input_cell = new TreeNode(cellInRange.Address);
                                nodes.Add(input_cell);
                            }

                            //Update the dependencies
                            range.addParent(input_cell);
                            input_cell.addChild(range);
                        }
                    }

                    MatchCollection matchedCells = Regex.Matches(formula, patternCell);  //matchedCells is a collection of all the cells that are referenced by the formula
                    foreach (Match m in matchedCells)
                    {
                        //TODO Currently influences between different worksheets do not work; this should be fixed
                        TreeNode input_cell = null;
                        //Find the node object for the current cell in the list of TreeNodes
                        foreach (TreeNode node in nodes)
                        {
                            if (node.getName().Replace("$", "") == m.Value.Replace("$", ""))
                            {
                                input_cell = node;
                            }
                            else
                            {
                                continue;
                            }
                        }

                        //Update the dependencies
                        n.addParent(input_cell);
                        input_cell.addChild(n);
                    }
                }
            }

            //Propagate weights
            foreach (TreeNode node in nodes)
            {
                if (!node.hasChildren())
                {
                    node.setWeight(1.0);  //Set the weight of all output nodes (and charts) to 1.0 to start
                    //Now we propagate proportional weights to all of this node's inputs
                    propagateWeight(node, 1.0);
                }
            }
            double max_weight = 0.0;  //Keep track of the max weight for normalizing later (used for coloring cells based on weight)
            foreach (TreeNode node in nodes)
            {
                if (node.getWeight() > max_weight)
                    max_weight = node.getWeight();
            }
            //TODO -- we are not able to capture ranges that are identified in stored procedures or macros, just ones referenced in formulas
            List<double> starting_outputs = new List<double>(); //This will store all the output nodes at the start of the procedure for swapping values
            List<TreeNode> output_cells = new List<TreeNode>();
            //Store all the starting output values
            foreach (TreeNode node in nodes)
            {
                if (!node.hasChildren() && !node.isChart())
                {
                    output_cells.Add(node);
                }
            }
            foreach (TreeNode n in output_cells)
            {
                starting_outputs.Add(activeWorksheet.get_Range(n.getName()).Value);
            }

            //Procedure for swapping values within ranges, one cell at a time
            if (!checkBox2.Checked) //Checks if the option for only analyzing the selection is checked
            {
                foreach (TreeNode node in nodes)
                {
                    //For every range node
                    if (node.isRange())
                    {
                        double[] influences = new double[node.getParents().Count]; //Array to keep track of the influence values for every cell
                        int influence_index = 0;        //Keeps track of the current position in the influences array
                        double max_total_delta = 0;     //The maximum influence found (for normalizing)
                        double min_total_delta = 0;     //The minimum influence found (for normalizing)
                        //Swapping values; loop over all nodes in the range
                        foreach (TreeNode parent in node.getParents())
                        {
                            Excel.Range cell = activeWorksheet.get_Range(parent.getName());
                            string formula = "";
                            if (cell.HasFormula)
                                formula = cell.Formula;
                            double start_value = cell.Value;
                            double total_delta = 0;
                            double delta = 0;
                            //Swapping loop - swap every sibling
                            foreach (TreeNode sibling in node.getParents())
                            {
                                if (sibling.getName() == parent.getName())
                                {
                                    continue;
                                }
                                Excel.Range sibling_cell = activeWorksheet.get_Range(sibling.getName());
                                cell.Value = sibling_cell.Value;
                                int index = 0;
                                delta = 0;
                                foreach (TreeNode n in output_cells)
                                {
                                    delta = Math.Abs(starting_outputs[index] - activeWorksheet.get_Range(n.getName()).Value) / starting_outputs[index];
                                    index++;
                                    total_delta = total_delta + delta;
                                }
                            }
                            total_delta = total_delta / (node.getParents().Count - 1);
                            influences[influence_index] = total_delta;
                            influence_index++;
                            //MessageBox.Show(cell.get_Address() + " Total delta = " + (total_delta * 100) + "%");
                            if (max_total_delta < total_delta)
                            {
                                max_total_delta = total_delta;
                            }
                            if (min_total_delta > total_delta || min_total_delta == 0)
                            {
                                min_total_delta = total_delta;
                            }
                            cell.Value = start_value;
                            if (formula != "")
                                cell.Formula = formula;
                            cell.Interior.Color = System.Drawing.Color.Beige;
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
                            Excel.Range cell = activeWorksheet.get_Range(parent.getName());
                            cell.Interior.Color = System.Drawing.Color.FromArgb(Convert.ToInt32(255 - influences[indexer] * 255), 255, 255);
                            indexer++;
                        }
                    }
                }
            }

            //Procedure for swapping values within ranges, replacing all repeated values at once
            if (checkBox2.Checked) //Checks if the option for only analyzing the selection is checked
            {
                foreach (TreeNode node in nodes)
                {
                    //For each range node, do the following:
                    if (node.isRange())
                    {
                        double[] influences = new double[node.getParents().Count];  //Array to keep track of the influence values for every cell
                        int influence_index = 0;        //Keeps track of the current position in the influences array
                        double max_total_delta = 0;     //The maximum influence found (for normalizing)
                        double min_total_delta = 0;     //The minimum influence found (for normalizing)
                        //Swapping values; loop over all nodes in the range
                        foreach (TreeNode parent in node.getParents())
                        {
                            String twin_cells_string = parent.getName();
                            //Find any nodes with a matching value and keep track of them
                            int twin_count = 1;     //This will keep track of the number of cells that have this exact value
                            foreach (TreeNode twin in node.getParents())
                            {
                                if (twin.getName() == parent.getName())
                                {
                                    continue;
                                }
                                if (activeWorksheet.get_Range(twin.getName()).Value == activeWorksheet.get_Range(parent.getName()).Value)
                                {
                                    twin_cells_string = twin_cells_string + "," + twin.getName();
                                    twin_count++;
                                }
                            }
                            //MessageBox.Show("Twin count: " + twin_count);
                            Excel.Range twin_cells = activeWorksheet.get_Range(twin_cells_string);
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
                            double start_value = activeWorksheet.get_Range(parent.getName()).Value;
                            double total_delta = 0;
                            double delta = 0;
                            foreach (TreeNode sibling in node.getParents())
                            {
                                if (sibling.getName() == parent.getName())
                                {
                                    continue;
                                }
                                Excel.Range sibling_cell = activeWorksheet.get_Range(sibling.getName());
                                twin_cells.Value = sibling_cell.Value;
                                int index = 0;
                                delta = 0;
                                foreach (TreeNode n in output_cells)
                                {
                                    delta = Math.Abs(starting_outputs[index] - activeWorksheet.get_Range(n.getName()).Value) / starting_outputs[index];
                                    index++;
                                    total_delta = total_delta + delta;
                                    //MessageBox.Show("Substituting " + sibling.getName() 
                                    //  + "\nDelta = |" + starting_outputs[index] + " - " + activeWorksheet.get_Range(n.getName()).Value + "| / " + starting_outputs[index]
                                    //  + " = " + delta
                                    //  + "\nTotal Delta = " + total_delta);
                                }
                            }
                            total_delta = total_delta / (node.getParents().Count - 1);
                            influences[influence_index] = total_delta / twin_count;
                            influence_index++;
                            //MessageBox.Show(twin_cells.get_Address() + " Total delta = " + (total_delta * 100) + "%");
                            if (max_total_delta < total_delta)
                            {
                                max_total_delta = total_delta;
                            }
                            if (min_total_delta > total_delta || min_total_delta == 0)
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
                            Excel.Range cell = activeWorksheet.get_Range(parent.getName());
                            cell.Interior.Color = System.Drawing.Color.FromArgb(Convert.ToInt32(255 - influences[indexer] * 255), 255, 255);
                            indexer++;
                        }
                    }
                }
            }


            //Print out text for GraphViz representation of the dependence graph
            string tree = "";
            foreach (TreeNode node in nodes)
            {
                tree += node.toGVString(max_weight) + "\n";
            }
            Display disp = new Display();
            disp.textBox1.Text = "digraph g{" + tree + "}";
            disp.ShowDialog();
        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            //IdentifyRanges();
            constructTree();
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
    }
}
