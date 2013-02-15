using System;
using System.Collections;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
using System.Text.RegularExpressions;

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
        public static TreeNode[][][] CreateFormulaNodes(ArrayList rs, Excel.Application app)
        {
            TreeNode[][][] nodes_grid;   //This is a multi-dimensional array of TreeNodes that will hold all the TreeNodes -- stores the dependence graph
            Excel.Workbook wb = app.ActiveWorkbook;

            // init nodes_grid
            nodes_grid = new TreeNode[app.Worksheets.Count + app.Charts.Count][][];
            int index = 0;
            foreach (Excel.Worksheet worksheet in app.Worksheets)
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

            foreach (Excel.Range worksheet_range in rs)
            {
                // Go through every cell of every worksheet
                if (worksheet_range != null)
                {
                    foreach (Excel.Range cell in worksheet_range)
                    {
                        if (cell.Value != null)
                        {
                            TreeNode n = new TreeNode(cell.Address, cell.Worksheet, wb);  //Create a TreeNode for every cell with the name being the cell's address and set the node's worksheet appropriately
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
            }
            return nodes_grid;
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

        public static void FindRangeReferencesWithQuotes(string formula, string worksheet_name, MatchCollection matchedRanges, Regex[] regex_array, int ws_index, System.Collections.Generic.List<TreeNode> ranges, TreeNode formula_cell, Excel.Worksheet ws_ref, Excel.Workbook activeWorkbook, Excel.Worksheet referencedWorksheet, TreeNode[][][] nodes_grid)
        {
            //First look for range references of the form 'worksheet_name'!A1:A10 in the formula (with quotation marks around the name)
            matchedRanges = regex_array[4 * (ws_index - 1)].Matches(formula);
            
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
                        //System.Windows.Forms.MessageBox.Show("Found node:" + ws_name + "_" + endCells[0] + ":" + endCells[1]);
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
                    range = new TreeNode(endCells[0].Replace("$", "") + "_to_" + endCells[1].Replace("$", ""), ws_ref, activeWorkbook);
                    //System.Windows.Forms.MessageBox.Show("Created range node:" + ws_name + "_" + endCells[0] + ":" + endCells[1]);
                    ranges.Add(range);
                }
                formula_cell.addParent(range);
                range.addChild(formula_cell);
                //Add each cell contained in the range to the dependencies
                foreach (Excel.Range cellInRange in referencedWorksheet.Range[endCells[0], endCells[1]])
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
                        input_cell = new TreeNode(cellInRange.Address, cellInRange.Worksheet, activeWorkbook);
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
        }

        public static void FindRangeReferencesWithoutQuotes(string formula, string worksheet_name, MatchCollection matchedRanges, Regex[] regex_array, int ws_index, System.Collections.Generic.List<TreeNode> ranges, TreeNode formula_cell, Excel.Worksheet ws_ref, Excel.Workbook activeWorkbook, Excel.Worksheet referencedWorksheet, TreeNode[][][] nodes_grid)
        {
            //Next look for range references of the form worksheet_name!A1:A10 in the formula (no quotation marks around the name)
            matchedRanges = regex_array[4 * (ws_index - 1) + 1].Matches(formula);
            
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
                        //System.Windows.Forms.MessageBox.Show("Found node:" + ws_name + "_" + endCells[0] + ":" + endCells[1]);
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
                    range = new TreeNode(endCells[0].Replace("$", "") + "_to_" + endCells[1].Replace("$", ""), ws_ref, activeWorkbook);
                    //System.Windows.Forms.MessageBox.Show("Created node:" + ws_name + "_" + endCells[0] + ":" + endCells[1]);
                    ranges.Add(range);
                }
                formula_cell.addParent(range);
                range.addChild(formula_cell);
                //Add each cell contained in the range to the dependencies
                foreach (Excel.Range cellInRange in referencedWorksheet.Range[endCells[0], endCells[1]])
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
                        input_cell = new TreeNode(cellInRange.Address, cellInRange.Worksheet, activeWorkbook);
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
        }

        public static void FindCellReferencesWithQuotes(string formula, string worksheet_name, MatchCollection matchedCells, Regex[] regex_array, int ws_index, System.Collections.Generic.List<TreeNode> ranges, TreeNode formula_cell, Excel.Worksheet ws_ref, Excel.Workbook activeWorkbook, Excel.Sheets worksheets, TreeNode[][][] nodes_grid)
        {
            // Now we look for references of the kind 'worksheet_name'!A1 (with quotation marks)
            matchedCells = regex_array[4 * (ws_index - 1) + 2].Matches(formula);
            
            foreach (Match match in matchedCells)
            {
                formula = formula.Replace(match.Value, "");
                string ws_name = worksheet_name; // match.Value.Substring(1, match.Value.LastIndexOf("!") - 2); // Get the name of the worksheet being referenced
                string cell_coordinates = match.Value.Substring(match.Value.LastIndexOf("!") + 1);
                //Get the actual cell that is being referenced
                Excel.Range input = null;
                foreach (Excel.Worksheet ws in worksheets)
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
                    input_cell = new TreeNode(cell_coordinates.Replace("$", ""), ws_ref, activeWorkbook);
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

        public static void FindCellReferencesWithoutQuotes(string formula, string worksheet_name, MatchCollection matchedCells, Regex[] regex_array, int ws_index, System.Collections.Generic.List<TreeNode> ranges, TreeNode formula_cell, Excel.Worksheet ws_ref, Excel.Workbook activeWorkbook, Excel.Sheets worksheets, TreeNode[][][] nodes_grid)
        {
            //Lastly we look for references of the kind worksheet_name!A1 (without quotation marks)
            matchedCells = regex_array[4 * (ws_index - 1) + 3].Matches(formula);
            
            foreach (Match match in matchedCells)
            {
                formula = formula.Replace(match.Value, "");
                string ws_name = worksheet_name; //match.Value.Substring(0, match.Value.LastIndexOf("!")); // Get the name of the worksheet being referenced
                string cell_coordinates = match.Value.Substring(match.Value.LastIndexOf("!") + 1);
                //System.Windows.Forms.MessageBox.Show(formula_cell.getName() + " refers to the cell " + ws_name + "!" + cell_coordinates);
                //Get the actual cell that is being referenced
                Excel.Range input = null;
                foreach (Excel.Worksheet ws in worksheets)
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
                    input_cell = new TreeNode(cell_coordinates.Replace("$", ""), ws_ref, activeWorkbook);
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

        public static void FindRangeReferencesInCurrentWorksheet(string formula, MatchCollection matchedRanges, MatchCollection matchedCells, Regex[] regex_array, int ws_index, System.Collections.Generic.List<TreeNode> ranges, TreeNode formula_cell, Excel.Workbook activeWorkbook, Excel.Sheets worksheets, TreeNode[][][] nodes_grid, Excel.Range c)
        {
            //First look for range references in the formula
            matchedRanges = regex_array[regex_array.Length - 2].Matches(formula);

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
                    //System.Windows.Forms.MessageBox.Show("Created range node:" + c.Worksheet.Name + "_" + endCells[0] + ":" + endCells[1]);
                    range = new TreeNode(endCells[0].Replace("$", "") + "_to_" + endCells[1].Replace("$", ""), c.Worksheet, activeWorkbook);
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
                        input_cell = new TreeNode(cellInRange.Address, cellInRange.Worksheet, activeWorkbook);
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
        }

        public static void FindNamedRangeReferences(string formula, MatchCollection matchedRanges, MatchCollection matchedCells, Regex[] regex_array, int ws_index, System.Collections.Generic.List<TreeNode> ranges, TreeNode formula_cell, Excel.Workbook activeWorkbook, Excel.Sheets worksheets, TreeNode[][][] nodes_grid, Excel.Range c, Excel.Names names)
        {
            //Find any references to named ranges
            //TODO -- this should probably be done in a better way - with a regular expression that will catch things like this:
            //"+range_name", "-range_name", "*range_name", etc., because right now a range name may be part of the name of a 
            //formula that is used. For instance a range could be named "s", and if the formula has the "sum" function in it, we will 
            //falsely detect a reference to "s". This does not affect the correctness of the algorithm, because all we care about 
            //from the dependence graph is identifying which cells are outputs, and identifying user-defined ranges
            //and this type of error will not affect either one
            foreach (Excel.Name named_range in names)
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
                        //System.Windows.Forms.MessageBox.Show("Created range node:" + c.Worksheet.Name + "_" + endCells[0] + ":" + endCells[1]);
                        range = new TreeNode(endCells[0].Replace("$", "") + "_to_" + endCells[1].Replace("$", ""), named_range.RefersToRange.Worksheet, activeWorkbook);
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
                            input_cell = new TreeNode(cellInRange.Address, cellInRange.Worksheet, activeWorkbook);
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
                    Excel.Range input = named_range.RefersToRange;
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
                        input_cell = new TreeNode(named_range.RefersToRange.Address.Replace("$", ""), named_range.RefersToRange.Worksheet, activeWorkbook);
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

        public static void FindCellReferencesInCurrentWorksheet(string formula, MatchCollection matchedRanges, MatchCollection matchedCells, Regex[] regex_array, int ws_index, System.Collections.Generic.List<TreeNode> ranges, TreeNode formula_cell, Excel.Workbook activeWorkbook, Excel.Sheets worksheets, TreeNode[][][] nodes_grid, Excel.Range c)
        {
            matchedCells = regex_array[regex_array.Length - 1].Matches(formula);
            foreach (Match m in matchedCells)
            {
                Excel.Range input = c.Worksheet.get_Range(m.Value);
                TreeNode input_cell = null;
                //System.Windows.Forms.MessageBox.Show(m.Value);
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
                    input_cell = new TreeNode(m.Value.Replace("$", ""), c.Worksheet, activeWorkbook);
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
