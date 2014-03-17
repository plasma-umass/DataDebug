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
using RangeDict = System.Collections.Generic.Dictionary<string, DataDebugMethods.TreeNode>;

using Microsoft.FSharp.Core;

namespace DataDebugMethods
{
    public static class ConstructTree
    {
        // This method constructs the dependency graph from the workbook.
        public static AnalysisData constructTree(Excel.Workbook wb, Excel.Application app, bool show_progbar)
        {
            // Make a new analysisData object
            AnalysisData data = new AnalysisData(app, app.ActiveWorkbook, !show_progbar);

            // Get a range representing the formula cells for each worksheet in each workbook
            ArrayList formulaRanges = ConstructTree.GetFormulaRanges(wb.Worksheets, app);

            // Create nodes for every cell containing a formula
            data.formula_nodes = ConstructTree.CreateFormulaNodes(formulaRanges, wb, app);

            //Now we parse the formulas in nodes to extract any range and cell references
            foreach(TreeDictPair pair in data.formula_nodes)
            {
                // This is a formula:
                TreeNode formula_node = pair.Value;

                // For each of the ranges found in the formula by the parser,
                // 1. make a new TreeNode for the range
                // 2. make TreeNodes for each of the cells in that range
                foreach (Excel.Range input_range in ExcelParserUtility.GetReferencesFromFormula(formula_node.getFormula(), formula_node.getWorkbookObject(), formula_node.getWorksheetObject()))
                {
                    // this function both creates a TreeNode and adds it to AnalysisData.input_ranges
                    TreeNode range_node = ConstructTree.MakeRangeTreeNode(data.input_ranges, input_range, formula_node);
                    // this function both creates cell TreeNodes for a range and adds it to AnalysisData.cell_nodes
                    ConstructTree.CreateCellNodesFromRange(range_node, formula_node, data.formula_nodes, data.cell_nodes, wb);
                }

                // For each single-cell input found in the formula by the parser,
                // link to output TreeNode if the input cell is a formula. This allows
                // us to consider functions with single-cell inputs as outputs.
                foreach (AST.Address input_addr in ExcelParserUtility.GetSingleCellReferencesFromFormula(formula_node.getFormula(), formula_node.getWorkbookObject(), formula_node.getWorksheetObject()))
                {
                    // Find the input cell's TreeNode; if there isn't one, move on.
                    // We don't care about scalar inputs that are not functions
                    TreeNode tn;
                    if (data.formula_nodes.TryGetValue(input_addr, out tn))
                    {
                        // sanity check-- should be a formula
                        if (tn.isFormula())
                        {
                            // link input to output formula node
                            tn.addOutput(formula_node);
                            formula_node.addInput(tn);
                        }
                    }
                }
            }
            return data;
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
        public static TreeDict CreateFormulaNodes(ArrayList rs, Excel.Workbook wb, Excel.Application app)
        {
            // init nodes
            var nodes = new TreeDict();

            foreach (Excel.Range worksheet_range in rs)
            {
                foreach (Excel.Range cell in worksheet_range)
                {
                    if (cell.Value2 != null)
                    {
                        var addr = AST.Address.AddressFromCOMObject(cell, wb);
                        var n = new TreeNode(cell, cell.Worksheet, wb);

                        if (cell.HasFormula)
                        {
                            //n.setIsFormula();   // I believe that this is unnecessary
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

        public static void CreateCellNodesFromRange(TreeNode input_range, TreeNode formula, TreeDict formula_nodes, TreeDict cell_nodes, Excel.Workbook wb)
        {
            foreach (Excel.Range cell in input_range.getCOMObject())
            {
                var addr = AST.Address.AddressFromCOMObject(cell, formula.getWorkbookObject());

                // cell might either be another formula or just a simple data cell;
                var d = cell.HasFormula ? formula_nodes : cell_nodes;

                // add to appropriate dictionary
                TreeNode cell_node;
                if (!d.TryGetValue(addr, out cell_node))
                {
                    cell_node = new TreeNode(cell, cell.Worksheet, formula.getWorkbookObject());
                    d.Add(addr, cell_node);
                }

                // Allow perturbation of every input_range that contains at least one value
                // TODO: fix; the Workbook reference here is not correct in the case of cross-workbook reference;
                // that said, having the wrong workbook doesn't actually have any bearing on the correctness of this call
                if ((cell.HasFormula && ExcelParserUtility.GetSCFormulaNames((string)cell.Formula, wb.FullName, cell.Worksheet, wb).Count() > 0)) //|| cell.Value2 != null)
                {
                    input_range.DontPerturb();
                }

                // link cell, range, and formula inputs and outputs together
                input_range.addInput(cell_node);
                cell_node.addOutput(formula);
                formula.addInput(cell_node);
            }
        }

        public static TreeNode MakeRangeTreeNode(RangeDict input_ranges, Excel.Range input_range, TreeNode parent)
        {
            // parse the address
            //var addr = AST.Address.AddressFromCOMObject(input_range, parent.getWorkbookObject());
            var addr = String.Intern(input_range.Address);

            // get it from dictionary, or, if it does not exist, create it, add to dict, and return new ref
            TreeNode tn;
            if (!input_ranges.TryGetValue(addr, out tn))
            {
                tn = new TreeNode(input_range, input_range.Worksheet, parent.getWorkbookObject());
                input_ranges.Add(addr, tn);
            }
            return tn;
        }

    } // ConstructTree class ends here
} // namespace ends here
