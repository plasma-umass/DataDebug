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
        public static void constructTree(AnalysisData analysisData, Excel.Workbook wb, Excel.Application app)
        {
            // Get a range representing the formula cells for each worksheet in each workbook
            ArrayList formulaRanges = ConstructTree.GetFormulaRanges(wb.Worksheets, app);

            // Create nodes for every cell containing a formula
            analysisData.formula_nodes = ConstructTree.CreateFormulaNodes(formulaRanges, wb, app);

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
                            tn.addOutput(formula_node);
                            formula_node.addInput(tn);
                        }
                    }
                }
            }
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
                        var addr = AST.Address.AddressFromCOMObject(cell, cell.Worksheet.Name, wb.Name, wb.FullName);
                        var n = new TreeNode(cell, cell.Worksheet, wb);
                        
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

        public static void CreateCellNodesFromRange(TreeNode rangeNode, TreeNode formulaNode, TreeDict formula_nodes, TreeDict cell_nodes)
        {
            foreach (Excel.Range cell in rangeNode.getCOMObject())
            {
                TreeNode cellNode = null;
                //See if there is an existing node for this cell already in nodes; if there is, do not add it again - just grab the existing one
                if (!formula_nodes.TryGetValue(ExcelParser.GetAddress(cell.Address[true, true, Excel.XlReferenceStyle.xlR1C1, false], formulaNode.getWorkbookObject(), cell.Worksheet), out cellNode))
                {
                    //TODO CORRECT THE WORKBOOK PARAMETER IN THIS LINE: (IT SHOULD BE THE WORKBOOK OF cell, WHICH SHOULD COME FROM GetReferencesFromFormula
                    var addr = AST.Address.AddressFromCOMObject(cell, cell.Worksheet.Name, formulaNode.getWorkbookObject().Name, formulaNode.getWorkbookObject().FullName);
                    cellNode = new TreeNode(cell, cell.Worksheet, formulaNode.getWorkbookObject());

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

                // Allow perturbation for every range which contains at least one concrete value (not a formula)
                if (!cell.HasFormula && cell.Value2 != null)
                {
                    rangeNode.Perturb();
                }

                // link cell, range, and formula inputs and outputs together
                rangeNode.addInput(cellNode);
                cellNode.addOutput(formulaNode);
                formulaNode.addInput(cellNode);
            }
        }

        public static TreeNode MakeRangeTreeNode(TreeDict input_ranges, Excel.Range input_range, TreeNode parent)
        {
            // parse the address
            var addr = AST.Address.AddressFromCOMObject(input_range,
                                                        input_range.Worksheet.Name,
                                                        parent.getWorkbookObject().Name,
                                                        parent.getWorkbookObject().FullName);

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
