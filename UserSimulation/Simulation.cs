using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using TreeNode = DataDebugMethods.TreeNode;
using CellDict = System.Collections.Generic.Dictionary<AST.Address, string>;
using TreeScore = System.Collections.Generic.Dictionary<DataDebugMethods.TreeNode, int>;
using TreeDict = System.Collections.Generic.Dictionary<AST.Address, DataDebugMethods.TreeNode>;
using System.Diagnostics;
using DataDebugMethods;

namespace UserSimulation
{
    public enum ErrorConditions
    {
        OK,
        ContainsNoInputs
    }

    public class Simulation
    {
        //CellDict saved_values = new CellDict();
        //HashSet<AST.Address> tool_highlights = new HashSet<AST.Address>();
        
        //IEnumerable<Tuple<double, TreeNode>> analysis_results = null;
        //AST.Address flagged_cell = null;
        ErrorConditions exit_state = ErrorConditions.OK;
        List<AST.Address> true_positives;
        List<AST.Address> false_positives;
        HashSet<AST.Address> false_negatives;

        // create and run a CheckCell simulation
        public Simulation(int nboots, string filename, double significance, CellDict errors, Excel.Application app)
        {
            // set of known good cells, initially empty
            HashSet<AST.Address> known_good = new HashSet<AST.Address>();

            // list of discovered errors, initially empty
            true_positives = new List<AST.Address>();
            false_positives = new List<AST.Address>();

            // open workbook
            Excel.Workbook wb = Utility.OpenWorkbook(filename, app);

            // build dependency graph
            var data = ConstructTree.constructTree(app.ActiveWorkbook, app, true);
            if (data.TerminalInputNodes().Length == 0)
            {
                exit_state = ErrorConditions.ContainsNoInputs;
                return;
            }

            // save original spreadsheet state
            CellDict original_inputs = SaveInputs(data.TerminalInputNodes(), wb);

            // force a recalculation before saving outputs, otherwise we may
            // erroneously conclude that the procedure did the wrong thing
            // based solely on Excel floating-point oddities
            InjectValues(app, wb, original_inputs);

            // save function outputs
            CellDict original_outputs = SaveOutputs(data.TerminalFormulaNodes(), wb);

            // inject errors
            InjectValues(app, wb, errors);

            // remove errors until none remain
            var errors_remain = true;
            while(errors_remain)
            {
                TreeScore scores;
                // Get bootstraps
                scores = Analysis.Bootstrap(nboots, data, app, true);

                // Compute quantiles based on user-supplied sensitivity
                var quantiles = Analysis.ComputeQuantile<int, TreeNode>(scores.Select(
                    pair => new Tuple<int, TreeNode>(pair.Value, pair.Key))
                );

                // Get top outlier
                var flagged_cell = Analysis.GetTopOutlier(quantiles, known_good, significance);
                if (flagged_cell == null)
                {
                    errors_remain = false;
                }
                else
                {
                    // check to see if the flagged value is actually an error
                    if (errors.ContainsKey(flagged_cell))
                    {
                        true_positives.Add(flagged_cell);
                    }
                    else
                    {
                        false_positives.Add(flagged_cell);
                    }

                    // correct flagged cell
                    flagged_cell.GetCOMObject(app).Value2 = original_inputs[flagged_cell];

                    // mark it as known good
                    known_good.Add(flagged_cell);
                }
            }

            // find all of the false negatives
            false_negatives = GetFalseNegatives(true_positives, false_positives, errors);

            // close workbook without saving
            wb.Close(false, "", false);
        }

        // return the set of false negatives
        public static HashSet<AST.Address> GetFalseNegatives(List<AST.Address> true_positives, List<AST.Address> false_positives, CellDict errors)
        {
            var fnset = new HashSet<AST.Address>();
            var tpset = new HashSet<AST.Address>(true_positives);
            var fpset = new HashSet<AST.Address>(false_positives);

            foreach(KeyValuePair<AST.Address, string> error in errors)
            {
                var addr = error.Key;
                if (!tpset.Contains(addr) && !fpset.Contains(addr))
                {
                    fnset.Add(addr);
                }
            }

            return fnset;
        }

        // save spreadsheet inputs to a CellDict
        public static CellDict SaveInputs(TreeNode[] input_ranges, Excel.Workbook wb)
        {
            var cd = new CellDict();
            foreach (TreeNode input_range in input_ranges)
            {
                foreach (TreeNode cell in input_range.getChildren())
                {
                    // never save formula; there's no point since we don't perturb them
                    var comcell = cell.getCOMObject();
                    if (!comcell.HasFormula)
                    {
                        cd.Add(cell.GetAddress(), cell.getCOMValueAsString());
                    }
                }
            }
            return cd;
        }

        // save spreadsheet outputs to a CellDict
        public static CellDict SaveOutputs(TreeNode[] formula_nodes, Excel.Workbook wb)
        {
            var cd = new CellDict();
            foreach (TreeNode formula_cell in formula_nodes)
            {
                // throw an exception in debug mode, because this should never happen
                Debug.Assert(formula_cell.getCOMObject().HasFormula);
                // save value
                cd.Add(formula_cell.GetAddress(), formula_cell.getCOMValueAsString());
            }
            return cd;
        }

        // inject errors into a workbook
        public static void InjectValues(Excel.Application app, Excel.Workbook wb, CellDict errors)
        {
            foreach (KeyValuePair<AST.Address, string> pair in errors)
            {
                var addr = pair.Key;
                var errorstr = pair.Value;
                var comcell = addr.GetCOMObject(app);

                // never perturb formulae
                if (!comcell.HasFormula)
                {
                    // inject error
                    addr.GetCOMObject(app).Value2 = errorstr;
                }
            }
        }
    }
}
