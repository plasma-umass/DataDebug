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
                        var addr = ExcelParser.GetAddress(cell.Address[true, true, Excel.XlReferenceStyle.xlR1C1, false], cell.Worksheet);
                        var n = new TreeNode(cell.Address, cell.Worksheet, wb);
                        
                        if (cell.HasFormula)
                        {
                            n.setIsFormula();
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

        public static void FindRangeReferencesWithQuotes(ref string formula, string worksheet_name, MatchCollection matchedRanges, Regex[] regex_array, int ws_index, TreeList ranges, TreeNode formula_cell, Excel.Worksheet ws_ref, Excel.Workbook activeWorkbook, Excel.Worksheet referencedWorksheet, TreeDict nodes)
        {
            //First look for range references of the form 'worksheet_name'!A1:A10 in the formula (with quotation marks around the name)
            matchedRanges = regex_array[4 * (ws_index - 1)].Matches(formula);
            
            foreach (Match match in matchedRanges)
            {
                formula = formula.Replace(match.Value, "");
                // Split up each matched range into the start and end cells of the range
                string range_coordinates = match.Value.Substring(match.Value.LastIndexOf("!") + 1);
                string[] endCells = range_coordinates.Split(':');
                string range_start = endCells[0];
                string range_end = endCells[1];

                // Try to find the range by name in existing TreeNodes
                TreeNode range = null;
                var range_name = range_start.Replace("$", "") + "_to_" + range_end.Replace("$", "");
                foreach (TreeNode n in ranges)
                {
                    if (n.getName().Replace("$", "") == range_name && n.getWorksheet() == worksheet_name)
                    {
                        range = n;
                    }
                }

                // If the range's TreeNode was not found, create it
                if (range == null)
                {
                    range = new TreeNode(range_name, ws_ref, activeWorkbook);
                    ranges.Add(range);
                }

                // Once we have a TreeNode for the range, we can update the parent-child relationship
                formula_cell.addParent(range);
                range.addChild(formula_cell);

                // Add each cell contained in the range to the dependencies
                foreach (Excel.Range cellInRange in referencedWorksheet.Range[range_start, range_end])
                {
                    TreeNode input_cell = null;
                    // Get the TreeNode if exists for this cell already, otherwise create it
                    if (!nodes.TryGetValue(Utility.ParseXLAddress(cellInRange), out input_cell))
                    {
                        // If it wasn't found, create a TreeNode for it
                        input_cell = new TreeNode(cellInRange.Address, cellInRange.Worksheet, activeWorkbook);
                    }

                    // Only add TreeNode to nodes if it is inside the UsedRange
                    if (Utility.InsideUsedRange(cellInRange))
                    {
                        nodes.Add(Utility.ParseXLAddress(cellInRange), input_cell);
                    }

                    // Update the dependencies, even if that means that input_cell is outside the UsedRange
                    // This is for diagnostic purposes
                    range.addParent(input_cell);
                    input_cell.addChild(range);
                }
            }
        }

        public static void FindRangeReferencesWithoutQuotes(ref string formula, string worksheet_name, MatchCollection matchedRanges, Regex[] regex_array, int ws_index, System.Collections.Generic.List<TreeNode> ranges, TreeNode formula_cell, Excel.Worksheet ws_ref, Excel.Workbook activeWorkbook, Excel.Worksheet referencedWorksheet, TreeDict nodes)
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
                    AST.Address addr = Utility.ParseXLAddress(cellInRange);
                    if (cellInRange.Column <= (cellInRange.Worksheet.UsedRange.Columns.Count + cellInRange.Worksheet.UsedRange.Column) && cellInRange.Row <= (cellInRange.Worksheet.UsedRange.Rows.Count + cellInRange.Worksheet.UsedRange.Row))
                    {
                        //if a TreeNode exists for this cell already
                        nodes.TryGetValue(addr, out input_cell);
                    }
                    //If it wasn't found, then it is blank, and we have to create a TreeNode for it
                    if (input_cell == null)
                    {
                        input_cell = new TreeNode(cellInRange.Address, cellInRange.Worksheet, activeWorkbook);
                        //Check if this cell's coordinates are within the bounds of the used range, otherwise there will be an index out of bounds error
                        if (cellInRange.Column <= (cellInRange.Worksheet.UsedRange.Columns.Count + cellInRange.Worksheet.UsedRange.Column) && cellInRange.Row <= (cellInRange.Worksheet.UsedRange.Rows.Count + cellInRange.Worksheet.UsedRange.Row))
                        {
                            nodes.Add(addr, input_cell);
                        }
                    }

                    //Update the dependencies
                    range.addParent(input_cell);
                    input_cell.addChild(range);
                }
            }
        }

        public static void FindCellReferencesWithQuotes(ref string formula, string worksheet_name, MatchCollection matchedCells, Regex[] regex_array, int ws_index, System.Collections.Generic.List<TreeNode> ranges, TreeNode formula_cell, Excel.Worksheet ws_ref, Excel.Workbook activeWorkbook, Excel.Sheets worksheets, TreeDict nodes)
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
                AST.Address addr = Utility.ParseXLAddress(input);

                if (input.Column <= (input.Worksheet.UsedRange.Columns.Count + input.Worksheet.UsedRange.Column) && input.Row <= (input.Worksheet.UsedRange.Rows.Count + input.Worksheet.UsedRange.Row))
                {
                    //if a TreeNode exists for this cell already, use it
                    nodes.TryGetValue(addr, out input_cell);
                }
                //If it wasn't found, then it is blank, and we have to create a TreeNode for it
                if (input_cell == null)
                {
                    input_cell = new TreeNode(cell_coordinates.Replace("$", ""), ws_ref, activeWorkbook);
                    //Check if this cell's coordinates are within the bounds of the used range, otherwise there will be an index out of bounds error
                    if (input.Column <= (input.Worksheet.UsedRange.Columns.Count + input.Worksheet.UsedRange.Column) && input.Row <= (input.Worksheet.UsedRange.Rows.Count + input.Worksheet.UsedRange.Row))
                    {
                        nodes.Add(addr, input_cell);
                    }
                }

                //Update the dependencies
                formula_cell.addParent(input_cell);
                input_cell.addChild(formula_cell);
            }
        }

        public static void FindCellReferencesWithoutQuotes(string formula, string worksheet_name, MatchCollection matchedCells, Regex[] regex_array, int ws_index, System.Collections.Generic.List<TreeNode> ranges, TreeNode formula_cell, Excel.Worksheet ws_ref, Excel.Workbook activeWorkbook, Excel.Sheets worksheets, TreeDict nodes)
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
                AST.Address addr = Utility.ParseXLAddress(input);
                
                if (input.Column <= (input.Worksheet.UsedRange.Columns.Count + input.Worksheet.UsedRange.Column) && input.Row <= (input.Worksheet.UsedRange.Rows.Count + input.Worksheet.UsedRange.Row))
                {
                    nodes.TryGetValue(addr, out input_cell);
                }
                //If it wasn't found, then it is blank, and we have to create a TreeNode for it
                if (input_cell == null)
                {
                    input_cell = new TreeNode(cell_coordinates.Replace("$", ""), ws_ref, activeWorkbook);
                    //Check if this cell's coordinates are within the bounds of the used range, otherwise there will be an index out of bounds error
                    if (input.Column <= (input.Worksheet.UsedRange.Columns.Count + input.Worksheet.UsedRange.Column) && input.Row <= (input.Worksheet.UsedRange.Rows.Count + input.Worksheet.UsedRange.Row))
                    {
                        nodes.Add(addr, input_cell);
                    }
                }

                //Update the dependencies
                formula_cell.addParent(input_cell);
                input_cell.addChild(formula_cell);
            }
        }

        public static void FindRangeReferencesInCurrentWorksheet(ref string formula, MatchCollection matchedRanges, MatchCollection matchedCells, Regex[] regex_array, int ws_index, System.Collections.Generic.List<TreeNode> ranges, TreeNode formula_cell, Excel.Workbook activeWorkbook, Excel.Sheets worksheets, TreeDict nodes, Excel.Range c)
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
                    AST.Address addr = Utility.ParseXLAddress(cellInRange);
                    

                    if (cellInRange.Column <= (cellInRange.Worksheet.UsedRange.Columns.Count + cellInRange.Worksheet.UsedRange.Row) && cellInRange.Row <= (cellInRange.Worksheet.UsedRange.Rows.Count + cellInRange.Worksheet.UsedRange.Row))
                    {
                        //if a TreeNode exists for this cell already, use it
                        nodes.TryGetValue(addr, out input_cell);
                    }
                    //If it wasn't found, then it is blank, and we have to create a TreeNode for it
                    if (input_cell == null)
                    {
                        input_cell = new TreeNode(cellInRange.Address, cellInRange.Worksheet, activeWorkbook);
                        //Check if this cell's coordinates are within the bounds of the used range, otherwise there will be an index out of bounds error
                        if (cellInRange.Column <= (cellInRange.Worksheet.UsedRange.Columns.Count + cellInRange.Worksheet.UsedRange.Column) && cellInRange.Row <= (cellInRange.Worksheet.UsedRange.Rows.Count + cellInRange.Worksheet.UsedRange.Row))
                        {
                            nodes.Add(addr, input_cell);
                        }
                    }

                    //Update the dependencies
                    range.addParent(input_cell);
                    input_cell.addChild(range);
                }
            }
        }

        public static void FindNamedRangeReferences(ref string formula, MatchCollection matchedRanges, MatchCollection matchedCells, Regex[] regex_array, int ws_index, System.Collections.Generic.List<TreeNode> ranges, TreeNode formula_cell, Excel.Workbook activeWorkbook, Excel.Sheets worksheets, TreeDict nodes, Excel.Range c, Excel.Names names)
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
                        AST.Address addr = Utility.ParseXLAddress(cellInRange);

                        if (cellInRange.Column <= (cellInRange.Worksheet.UsedRange.Columns.Count + cellInRange.Worksheet.UsedRange.Column) && cellInRange.Row <= (cellInRange.Worksheet.UsedRange.Rows.Count + cellInRange.Worksheet.UsedRange.Row))
                        {
                            //if a TreeNode exists for this cell already, use it
                            nodes.TryGetValue(addr, out input_cell);
                        }
                        //If it wasn't found, then it is blank, and we have to create a TreeNode for it
                        if (input_cell == null)
                        {
                            input_cell = new TreeNode(cellInRange.Address, cellInRange.Worksheet, activeWorkbook);
                            //Check if this cell's coordinates are within the bounds of the used range, otherwise there will be an index out of bounds error
                            if (cellInRange.Column <= (cellInRange.Worksheet.UsedRange.Columns.Count + cellInRange.Worksheet.UsedRange.Column) && cellInRange.Row <= (cellInRange.Worksheet.UsedRange.Rows.Count + cellInRange.Worksheet.UsedRange.Row))
                            {
                                nodes.Add(addr, input_cell);
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
                    AST.Address addr = Utility.ParseXLAddress(input);

                    if (input.Column <= (input.Worksheet.UsedRange.Columns.Count + input.Worksheet.UsedRange.Column) && input.Row <= (input.Worksheet.UsedRange.Rows.Count + input.Worksheet.UsedRange.Row))
                    {
                        //if a TreeNode exists for this cell already, use it
                        nodes.TryGetValue(addr, out input_cell);
                    }
                    //If it wasn't found, then it is blank, and we have to create a TreeNode for it
                    if (input_cell == null)
                    {
                        input_cell = new TreeNode(named_range.RefersToRange.Address.Replace("$", ""), named_range.RefersToRange.Worksheet, activeWorkbook);
                        //Check if this cell's coordinates are within the bounds of the used range, otherwise there will be an index out of bounds error
                        if (input.Column <= (input.Worksheet.UsedRange.Columns.Count + input.Worksheet.UsedRange.Column) && input.Row <= (input.Worksheet.UsedRange.Rows.Count + input.Worksheet.UsedRange.Row))
                        {
                            nodes.Add(addr, input_cell);
                        }
                    }
                    //Update the dependencies
                    formula_cell.addParent(input_cell);
                    input_cell.addChild(formula_cell);
                }
            }
        }

        public static void FindCellReferencesInCurrentWorksheet(ref string formula, MatchCollection matchedRanges, MatchCollection matchedCells, Regex[] regex_array, int ws_index, System.Collections.Generic.List<TreeNode> ranges, TreeNode formula_cell, Excel.Workbook activeWorkbook, Excel.Sheets worksheets, TreeDict nodes, Excel.Range c)
        {
            matchedCells = regex_array[regex_array.Length - 1].Matches(formula);
            foreach (Match m in matchedCells)
            {
                Excel.Range input = c.Worksheet.get_Range(m.Value);
                TreeNode input_cell = null;
                //System.Windows.Forms.MessageBox.Show(m.Value);
                //Find the node object for the current cell in the existing TreeNodes
                //Check if this cell's coordinates are within the bounds of the used range, otherwise there will be an index out of bounds error
                AST.Address addr = Utility.ParseXLAddress(input);
                
                if (input.Column <= (input.Worksheet.UsedRange.Columns.Count + input.Worksheet.UsedRange.Column) && input.Row <= (input.Worksheet.UsedRange.Rows.Count + input.Worksheet.UsedRange.Row))
                {
                    //if a TreeNode exists for this cell already, use it
                    nodes.TryGetValue(addr, out input_cell);
                }
                //If it wasn't found, then it is blank, and we have to create a TreeNode for it
                if (input_cell == null)
                {
                    input_cell = new TreeNode(m.Value.Replace("$", ""), c.Worksheet, activeWorkbook);
                    //Check if this cell's coordinates are within the bounds of the used range, otherwise there will be an index out of bounds error
                    if (input.Column <= (input.Worksheet.UsedRange.Columns.Count + input.Worksheet.UsedRange.Column) && input.Row <= (input.Worksheet.UsedRange.Rows.Count + input.Worksheet.UsedRange.Row))
                    {
                        nodes.Add(addr, input_cell);
                    }
                }
                //Update the dependencies
                formula_cell.addParent(input_cell);
                input_cell.addChild(formula_cell);
            }
        }

        public static void FindReferencesInCharts(Regex[] regex_array, System.Collections.Generic.List<TreeNode> ranges, Excel.Workbook activeWorkbook, Excel.Sheets charts, TreeDict nodes, string[] worksheet_names, Excel.Worksheet[] worksheet_refs, Excel.Sheets worksheets, Excel.Names names,TreeList nodelist)
        {
            foreach (Excel.Chart chart in charts)
            {
                //TODO The naming convention for TreeNode charts is kind of a hack; could fail if two charts have the same names when white spaces are removed - maybe add a random hash at the end
                TreeNode chart_node = new TreeNode(chart.Name, null, activeWorkbook);
                chart_node.setChart(true);
                charts.Add(chart_node);
                foreach (Excel.Series series in (Excel.SeriesCollection)chart.SeriesCollection(Type.Missing))
                {
                    string formula = series.Formula;  //The formula contained in the cell

                    MatchCollection matchedRanges = null;
                    MatchCollection matchedCells = null;
                    int ws_index = 1;
                    //foreach (string s in worksheet_names)
                    for (int i = 0; i < worksheet_names.Count(); i++)
                    {
                        string s = worksheet_names[i];
                        Excel.Worksheet ws_ref = worksheet_refs[i];
                        string worksheet_name = s.Replace("+", @"\+").Replace("^", @"\^").Replace("$", @"\$").Replace(".", @"\."); //Escape certain characters in the regular expression
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
                            chart_node.addParent(range);
                            range.addChild(chart_node);
                            //Add each cell contained in the range to the dependencies
                            foreach (Excel.Range cellInRange in worksheets[ws_index].Range[endCells[0], endCells[1]])
                            {
                                TreeNode input_cell = null;
                                //Find the node object for the current cell in the existing TreeNodes
                                //Check if this cell's coordinates are within the bounds of the used range, otherwise there will be an index out of bounds error
                                AST.Address addr = Utility.ParseXLAddress(cellInRange);
                                if (cellInRange.Column <= (cellInRange.Worksheet.UsedRange.Columns.Count + cellInRange.Worksheet.UsedRange.Column) && cellInRange.Row <= (cellInRange.Worksheet.UsedRange.Rows.Count + cellInRange.Worksheet.UsedRange.Row))
                                {
                                    //if a TreeNode exists for this cell already
                                    nodes.TryGetValue(addr, out input_cell);
                                }
                                //If it wasn't found, then it is blank, and we have to create a TreeNode for it
                                if (input_cell == null)
                                {
                                    input_cell = new TreeNode(cellInRange.Address, cellInRange.Worksheet, activeWorkbook);
                                    //Check if this cell's coordinates are within the bounds of the used range, otherwise there will be an index out of bounds error
                                    if (cellInRange.Column <= (cellInRange.Worksheet.UsedRange.Columns.Count + cellInRange.Worksheet.UsedRange.Column) && cellInRange.Row <= (cellInRange.Worksheet.UsedRange.Rows.Count + cellInRange.Worksheet.UsedRange.Row))
                                    {
                                        nodes.Add(addr, input_cell);
                                    }
                                }

                                //Update the dependencies
                                range.addParent(input_cell);
                                input_cell.addChild(range);
                            }
                        }

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
                            //If it was not found, create it
                            if (range == null)
                            {
                                range = new TreeNode(endCells[0].Replace("$", "") + "_to_" + endCells[1].Replace("$", ""), ws_ref, activeWorkbook);
                                //System.Windows.Forms.MessageBox.Show("Created range node:" + ws_name + "_" + endCells[0] + ":" + endCells[1]);
                                ranges.Add(range);
                            }
                            chart_node.addParent(range);
                            range.addChild(chart_node);
                            //Add each cell contained in the range to the dependencies
                            foreach (Excel.Range cellInRange in worksheets[ws_index].Range[endCells[0], endCells[1]])
                            {
                                TreeNode input_cell = null;
                                //Find the node object for the current cell in the existing TreeNodes
                                //Check if this cell's coordinates are within the bounds of the used range, otherwise there will be an index out of bounds error
                                AST.Address addr = Utility.ParseXLAddress(cellInRange);
                                if (cellInRange.Column <= (cellInRange.Worksheet.UsedRange.Columns.Count + cellInRange.Worksheet.UsedRange.Column) && cellInRange.Row <= (cellInRange.Worksheet.UsedRange.Rows.Count + cellInRange.Worksheet.UsedRange.Row))
                                {
                                    //if a TreeNode exists for this cell already
                                    nodes.TryGetValue(addr, out input_cell);
                                }
                                //If it wasn't found, then it is blank, and we have to create a TreeNode for it
                                if (input_cell == null)
                                {
                                    input_cell = new TreeNode(cellInRange.Address, cellInRange.Worksheet, activeWorkbook);
                                    //Check if this cell's coordinates are within the bounds of the used range, otherwise there will be an index out of bounds error
                                    if (cellInRange.Column <= (cellInRange.Worksheet.UsedRange.Columns.Count + cellInRange.Worksheet.UsedRange.Column) && cellInRange.Row <= (cellInRange.Worksheet.UsedRange.Rows.Count + cellInRange.Worksheet.UsedRange.Row))
                                    {
                                        nodes.Add(addr, input_cell);
                                    }
                                }

                                //Update the dependencies
                                range.addParent(input_cell);
                                input_cell.addChild(range);
                            }
                        }

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
                            AST.Address addr = Utility.ParseXLAddress(input);

                            if (input.Column <= (input.Worksheet.UsedRange.Columns.Count + input.Worksheet.UsedRange.Column) && input.Row <= (input.Worksheet.UsedRange.Rows.Count + input.Worksheet.UsedRange.Row))
                            {
                                //if a TreeNode exists for this cell already, use it
                                nodes.TryGetValue(addr, out input_cell);
                            }
                            //If it wasn't found, then it is blank, and we have to create a TreeNode for it
                            if (input_cell == null)
                            {
                                input_cell = new TreeNode(cell_coordinates.Replace("$", ""), ws_ref, activeWorkbook);
                                //Check if this cell's coordinates are within the bounds of the used range, otherwise there will be an index out of bounds error
                                if (input.Column <= (input.Worksheet.UsedRange.Columns.Count + input.Worksheet.UsedRange.Column) && input.Row <= (input.Worksheet.UsedRange.Rows.Count + input.Worksheet.UsedRange.Row))
                                {
                                    nodes.Add(addr, input_cell);
                                }
                            }

                            //Update the dependencies
                            chart_node.addParent(input_cell);
                            input_cell.addChild(chart_node);
                        }

                        //Lastly we look for references of the kind worksheet_name!A1 (without quotation marks)
                        matchedCells = regex_array[4 * (ws_index - 1) + 3].Matches(formula);
                        foreach (Match match in matchedCells)
                        {
                            formula = formula.Replace(match.Value, "");
                            string ws_name = worksheet_name; // match.Value.Substring(0, match.Value.LastIndexOf("!")); // Get the name of the worksheet being referenced
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
                            AST.Address addr = Utility.ParseXLAddress(input);

                            if (input.Column <= (input.Worksheet.UsedRange.Columns.Count + input.Worksheet.UsedRange.Column) && input.Row <= (input.Worksheet.UsedRange.Rows.Count + input.Worksheet.UsedRange.Row))
                            {
                                //if a TreeNode exists for this cell already, use it
                                nodes.TryGetValue(addr, out input_cell);
                            }
                            //If it wasn't found, then it is blank, and we have to create a TreeNode for it
                            if (input_cell == null)
                            {
                                input_cell = new TreeNode(cell_coordinates.Replace("$", ""), ws_ref, activeWorkbook);
                                //Check if this cell's coordinates are within the bounds of the used range, otherwise there will be an index out of bounds error
                                if (input.Column <= (input.Worksheet.UsedRange.Columns.Count + input.Worksheet.UsedRange.Column) && input.Row <= (input.Worksheet.UsedRange.Rows.Count + input.Worksheet.UsedRange.Row))
                                {
                                    nodes.Add(addr, input_cell);
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
                    foreach (Excel.Name named_range in names)
                    {
                        if (formula.Contains(named_range.Name))
                        {
                            formula = formula.Replace(named_range.Name, "");
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
                        }
                        //If it does not exist, create it
                        if (range == null)
                        {
                            range = new TreeNode(endCells[0].Replace("$", "") + "_to_" + endCells[1].Replace("$", ""), named_range.RefersToRange.Worksheet, activeWorkbook);
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
                            AST.Address addr = Utility.ParseXLAddress(cellInRange);

                            if (cellInRange.Column <= (cellInRange.Worksheet.UsedRange.Columns.Count + cellInRange.Worksheet.UsedRange.Column) && cellInRange.Row <= (cellInRange.Worksheet.UsedRange.Rows.Count + cellInRange.Worksheet.UsedRange.Row))
                            {
                                //if a TreeNode exists for this cell already, use it
                                nodes.TryGetValue(addr, out input_cell);
                            }
                            //If it wasn't found, then it is blank, and we have to create a TreeNode for it
                            if (input_cell == null)
                            {
                                input_cell = new TreeNode(cellInRange.Address, cellInRange.Worksheet, activeWorkbook);
                                //Check if this cell's coordinates are within the bounds of the used range, otherwise there will be an index out of bounds error
                                if (cellInRange.Column <= (cellInRange.Worksheet.UsedRange.Columns.Count + cellInRange.Worksheet.UsedRange.Column) && cellInRange.Row <= (cellInRange.Worksheet.UsedRange.Rows.Count + cellInRange.Worksheet.UsedRange.Row))
                                {
                                    nodes.Add(addr, input_cell);
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

            foreach (Excel.Worksheet worksheet in worksheets)
            {
                foreach (Excel.ChartObject chart in (Excel.ChartObjects)worksheet.ChartObjects(Type.Missing))
                {
                    //TODO The naming convention for TreeNode charts is kind of a hack; could fail if two charts have the same names when white spaces are removed - maybe add a random hash at the end
                    TreeNode chart_node = new TreeNode(chart.Name, worksheet, activeWorkbook);
                    chart_node.setChart(true);
                    nodelist.Add(chart_node);
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
                        //foreach (string s in worksheet_names)
                        for (int i = 0; i < worksheet_names.Count(); i++)
                        {
                            string s = worksheet_names[i];
                            Excel.Worksheet ws_ref = worksheet_refs[i];
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
                                foreach (TreeNode n in nodelist)
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
                                    //System.Windows.Forms.MessageBox.Show("Created range node:" + ws_name + "_" + endCells[0] + ":" + endCells[1]);
                                    nodelist.Add(range);
                                }
                                chart_node.addParent(range);
                                range.addChild(chart_node);
                                //Add each cell contained in the range to the dependencies
                                foreach (Excel.Range cellInRange in worksheets[ws_index].Range[endCells[0], endCells[1]])
                                {
                                    TreeNode input_cell = null;
                                    //Find the node object for the current cell in the list of TreeNodes
                                    foreach (TreeNode node in nodelist)
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
                                        input_cell = new TreeNode(cellInRange.Address, cellInRange.Worksheet, activeWorkbook);
                                        nodelist.Add(input_cell);
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
                                foreach (TreeNode n in nodelist)
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
                                    nodelist.Add(range);
                                }

                                //Update the dependencies
                                chart_node.addParent(range);
                                range.addChild(chart_node);
                                //Add each cell contained in the range to the dependencies
                                foreach (Excel.Range cellInRange in worksheets[ws_index].Range[endCells[0], endCells[1]])
                                {
                                    TreeNode input_cell = null;
                                    //Find the node object for the current cell in the list of TreeNodes
                                    foreach (TreeNode node in nodelist)
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
                                        input_cell = new TreeNode(cellInRange.Address, cellInRange.Worksheet, activeWorkbook);
                                        nodelist.Add(input_cell);
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
                                foreach (TreeNode node in nodelist)
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
                                    input_cell = new TreeNode(cell_coordinates.Replace("$", ""), ws_ref, activeWorkbook);
                                    nodelist.Add(input_cell);
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
                                foreach (TreeNode node in nodelist)
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
                                    input_cell = new TreeNode(cell_coordinates.Replace("$", ""), ws_ref, activeWorkbook);
                                    nodelist.Add(input_cell);
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
                        System.Collections.Generic.List<Excel.Range> rangeList = new System.Collections.Generic.List<Excel.Range>();
                        foreach (Match match in matchedRanges)
                        {
                            formula = formula.Replace(match.Value, "");
                            string[] endCells = match.Value.Split(':');     //Split up each matched range into the start and end cells of the range
                            TreeNode range = null;
                            //Try to find the range in existing TreeNodes
                            foreach (TreeNode node in nodelist)
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
                                range = new TreeNode(endCells[0].Replace("$", "") + "_to_" + endCells[1].Replace("$", ""), worksheet, activeWorkbook);
                                nodelist.Add(range);
                            }

                            //Update the dependencies
                            chart_node.addParent(range);
                            range.addChild(chart_node);
                            //Add each cell contained in the range to the dependencies
                            foreach (Excel.Range cellInRange in worksheet.Range[endCells[0], endCells[1]])
                            {
                                TreeNode input_cell = null;
                                //Find the node object for the current cell in the list of TreeNodes
                                foreach (TreeNode node in nodelist)
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
                                    input_cell = new TreeNode(cellInRange.Address, cellInRange.Worksheet, activeWorkbook);
                                    nodelist.Add(input_cell);
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
                        foreach (Excel.Name named_range in names)
                        {
                            if (formula.Contains(named_range.Name))
                            {
                                formula = formula.Replace(named_range.Name, "");
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
                            }
                            //If it does not exist, create it
                            if (range == null)
                            {
                                //System.Windows.Forms.MessageBox.Show("Created range node:" + c.Worksheet.Name + "_" + endCells[0] + ":" + endCells[1]);
                                range = new TreeNode(endCells[0].Replace("$", "") + "_to_" + endCells[1].Replace("$", ""), named_range.RefersToRange.Worksheet, activeWorkbook);
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
                                AST.Address addr = Utility.ParseXLAddress(cellInRange);

                                if (cellInRange.Column <= (cellInRange.Worksheet.UsedRange.Columns.Count + cellInRange.Worksheet.UsedRange.Column) && cellInRange.Row <= (cellInRange.Worksheet.UsedRange.Rows.Count + cellInRange.Worksheet.UsedRange.Row))
                                {
                                    //if a TreeNode exists for this cell already, use it
                                    nodes.TryGetValue(addr, out input_cell);
                                }
                                //If it wasn't found, then it is blank, and we have to create a TreeNode for it
                                if (input_cell == null)
                                {
                                    input_cell = new TreeNode(cellInRange.Address, cellInRange.Worksheet, activeWorkbook);
                                    //Check if this cell's coordinates are within the bounds of the used range, otherwise there will be an index out of bounds error
                                    if (cellInRange.Column <= (cellInRange.Worksheet.UsedRange.Columns.Count + cellInRange.Worksheet.UsedRange.Column) && cellInRange.Row <= (cellInRange.Worksheet.UsedRange.Rows.Count + cellInRange.Worksheet.UsedRange.Row))
                                    {
                                        nodes.Add(addr, input_cell);
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
                            foreach (TreeNode node in nodelist)
                            {
                                if (node.getName().Replace("$", "") == m.Value.Replace("$", "") && node.getWorksheet() == worksheet.Name)
                                {
                                    input_cell = node;
                                }
                            }
                            //If it wasn't found, then it is blank, and we have to create a TreeNode for it
                            if (input_cell == null)
                            {
                                input_cell = new TreeNode(m.Value, worksheet, activeWorkbook);
                                nodelist.Add(input_cell);
                            }
                            //Update the dependencies
                            chart_node.addParent(input_cell);
                            input_cell.addChild(chart_node);
                        }

                    }
                }
            }
        }

        public static void StoreOutputs(System.Collections.Generic.List<StartValue> starting_outputs, System.Collections.Generic.List<TreeNode> output_cells, TreeDict nodes)
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

        public static string GenerateGraphVizTree(System.Collections.Generic.List<TreeNode> nodes)
        {
            string tree = "";
            foreach (TreeNode node in nodes)
            {
                tree += node.toGVString(0) + "\n";
            }
            return "digraph g{" + tree + "}"; 
        }
    }
}
