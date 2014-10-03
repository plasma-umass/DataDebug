using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
using System.Text.RegularExpressions;
using Microsoft.FSharp.Core;

namespace DataDebugMethods
{
    public static class ConstructTree
    {
        public static AddressCache constructTree(Excel.Workbook wb, Excel.Application app, bool ignore_parse_errors)
        {
            return constructTree(wb, app, null, ignore_parse_errors);
        }

        // This method constructs the dependency graph from the workbook.
        public static AddressCache constructTree(Excel.Workbook wb, Excel.Application app, ProgBar pb, bool ignore_parse_errors)
        {
            //Start timing tree construction
            System.Diagnostics.Stopwatch sw = new System.Diagnostics.Stopwatch();
            sw.Start();

            // Use a fast array read to associate all cell references with their addresses
            var addrcache = new AddressCache(wb, app);

            // Parse formula nodes to extract references
            foreach(AST.Address formula_addr in addrcache.GetFormulaAddrs())
            {
                var formula_ref = addrcache.GetCOMObjectForAddress(formula_addr);

                // For each of the ranges found in the formula by the parser,
                // 1. make a new TreeNode for the range
                // 2. make TreeNodes for each of the cells in that range
                foreach (AST.Range input_range in ExcelParserUtility.GetReferencesFromFormula(formula_ref, ignore_parse_errors))
                {
                    // Fetch/create COMRef, adding it to the cache, if necessary
                    var range_ref = addrcache.MakeCOMRef(input_range, formula_ref);

                    // this function both creates cell TreeNodes for a range and adds it to AnalysisData.cell_nodes
                    ConstructTree.CreateCellNodesFromRange(range_node, formula_addr, data.formula_nodes, data.cell_nodes, wb, ignore_parse_errors);
                }

                // For each single-cell input found in the formula by the parser,
                // link to output TreeNode if the input cell is a formula. This allows
                // us to consider functions with single-cell inputs as outputs.
                IEnumerable<AST.Address> input_addrs;
                try
                {
                    input_addrs = ExcelParserUtility.GetSingleCellReferencesFromFormula(formula_addr.Formula, formula_addr.getWorkbookObject(), formula_addr.getWorksheetObject(), ignore_parse_errors);
                }
                catch (ExcelParserUtility.ParseException)
                {
                    // on parse exception, return an empty sequence
                    input_addrs = new List<AST.Address>();
                }
                foreach (AST.Address input_addr in input_addrs)
                {
                    // Find the input cell's TreeNode;
                    // Find out if it is a formula
                    TreeNode tn;
                    if (data.formula_nodes.TryGetValue(input_addr, out tn))
                    {
                        // sanity check-- should be a formula
                        if (tn.isFormula())
                        {
                            // link input to output formula node
                            tn.addOutput(formula_addr);
                            formula_addr.addInput(tn);
                        }
                    }
                    else   //If it's not a formula, then it is a scalar value; we create a TreeNode for it and put it in data.cell_nodes
                    {  
                        // if we have already created a TreeNode for this, add a connection to the formula that referenced it
                        if (data.cell_nodes.TryGetValue(input_addr, out tn))
                        {
                            tn.addOutput(formula_addr);
                            formula_addr.addInput(tn);
                        }
                        //otherwise create a new TreeNode and connect it up to the formula
                        else
                        {
                            Excel.Range cell = input_addr.GetCOMObject(app);
                            tn = new TreeNode(cell, cell.Worksheet, (Excel.Workbook)cell.Worksheet.Parent);
                            data.cell_nodes.Add(input_addr, tn);
                            tn.addOutput(formula_addr);
                            formula_addr.addInput(tn);
                        }
                    }
                }
            }

            //Stop timing the tree construction
            sw.Stop();
            //Store elapsed time
            TimeSpan elapsed = sw.Elapsed;
            data.tree_construct_time = elapsed.TotalSeconds;

            return data;
        }

        // This method returns an ArrayList of formula ranges, one range per worksheet
        public static ArrayList GetFormulaRanges(Excel.Sheets ws, Excel.Application app)
        {
            var fn_filter = new Regex("^=", RegexOptions.Compiled);

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
                    if (cell.HasFormula
                        && !String.IsNullOrWhiteSpace(System.Convert.ToString(cell.Formula))
                        && fn_filter.IsMatch(System.Convert.ToString(cell.Formula)))
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

        public static void CreateCellNodesFromRange(TreeNode input_range, TreeNode formula, TreeDict formula_nodes, TreeDict cell_nodes, Excel.Workbook wb, bool ignore_parse_errors)
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
                if ((cell.HasFormula && ExcelParserUtility.GetSCFormulaNames((string)cell.Formula, wb.FullName, cell.Worksheet, wb, ignore_parse_errors).Count() > 0)) //|| cell.Value2 != null)
                {
                    input_range.SetDoNotPerturb();
                }

                // link cell, range, and formula inputs and outputs together
                input_range.addInput(cell_node);
                cell_node.addOutput(formula);
                formula.addInput(cell_node);
            }
        }

        public static TreeNode MakeRangeTreeNode(RangeDict input_ranges, Excel.Range com_range, TreeNode parent)
        {
            // get COMRef
            

            // parse the absolute address
            var addr = String.Intern(com_range.get_Address(true, true));

            // get it from dictionary, or, if it does not exist, create it, add to dict, and return new ref
            TreeNode tn;
            if (!input_ranges.TryGetValue(addr, out tn))
            {
                tn = new TreeNode(com_range, com_range.Worksheet, parent.getWorkbookObject());
                input_ranges.Add(addr, tn);
            }
            return tn;
        }

    } // ConstructTree class ends here
} // namespace ends here
