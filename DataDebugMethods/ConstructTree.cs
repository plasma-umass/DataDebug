using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
using System.Text.RegularExpressions;

namespace DataDebugMethods
{
    public static class ConstructTree
    {
        public static int CountFormulaCells(Excel.Range[] rs)
        {
            int count = 0;
            foreach (var r in rs)
            {
                if (r != null)
                {
                    count += r.Cells.Count;
                }
            }
            return count;
        }

        public static Excel.Range[] GetFormulaRanges(Excel.Sheets ws, Excel.Application app)
        {
            Excel.Range[] analysisRanges = new Excel.Range[ws.Count]; //This keeps track of the range to be analyzed in every worksheet of the workbook

            int worksheet_index = 0; // keeps track of which worksheet we are currently examining
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
                    analysisRanges[worksheet_index] = formula_cells;
                }
                // we found no cells
                else
                {
                    analysisRanges[worksheet_index] = null;
                }
                // point at the next worksheet in analysisRanges
                worksheet_index++;
            }
            return analysisRanges;
        }

        //First we create nodes for every non-null cell; then we will operate on these node objects, connecting them in the tree, etc. 
        //This includes cells that contain constants and formulas
        //Go through every worksheet
        public static TreeNode[][][] CreateFormulaNodes(Excel.Range[] rs, Excel.Application app)
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
    }
}
