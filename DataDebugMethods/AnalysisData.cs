using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
using TreeDict = System.Collections.Generic.Dictionary<AST.Address, DataDebugMethods.TreeNode>;
using RangeDict = System.Collections.Generic.Dictionary<string, DataDebugMethods.TreeNode>;
using System.Diagnostics;

namespace DataDebugMethods
{
    public class AnalysisData
    {
        public List<TreeNode> nodelist;     // holds all the TreeNodes in the Excel file
        public RangeDict input_ranges;
        public TreeDict formula_nodes;
        public TreeDict cell_nodes;
        public Excel.Sheets charts;
        public double tree_construct_time;
        private ProgBar pb;

        private int _pb_max;
        private int _pb_count = 0;

        public AnalysisData(Excel.Application application, Excel.Workbook wb) 
        {
            charts = wb.Charts;
            nodelist = new List<TreeNode>();
            input_ranges = new RangeDict();
            cell_nodes = new TreeDict();
        }

        public AnalysisData(Excel.Application application, Excel.Workbook wb, ProgBar progbar)
            : this (application, wb)
        {
            pb = progbar;
        }

        public void SetProgress(int i)
        {
            if (pb != null) pb.SetProgress(i);
        }

        public void SetPBMax(int max)
        {
            _pb_max = max;
        }

        public void PokePB()
        {
            if (pb != null)
            {
                _pb_count += 1;
                this.SetProgress(_pb_count * 100 / _pb_max);
            }
        }

        private void KillPB()
        {
            // Kill progress bar
            if (pb != null) pb.Close();
        }

        public TreeNode[] TerminalFormulaNodes(bool all_outputs)
        {
            // return only the formula nodes which do not provide
            // input to any other cell and which are also not
            // in our list of excluded functions
            if (all_outputs)
            {
                return formula_nodes.Select(pair => pair.Value).ToArray();
            }
            else
            {
                return formula_nodes.Where(pair => pair.Value.getOutputs().Count == 0)
                                    .Select(pair => pair.Value).ToArray();
            }
        }

        public TreeNode[] TerminalInputNodes()
        {
            // this should filter out the following two cases:
            // 1. input range is intermediate (acts as input to a formula
            //    and also contains a formula which consumes input from
            //    another range).
            // 2. the range is actually a formula cell
            return input_ranges.Where(pair => !pair.Value.GetDontPerturb()
                                              && !pair.Value.isFormula())
                               .Select(pair => pair.Value).ToArray();
        }

        /// <summary>
        /// This method returns all input TreeNodes that are guaranteed to be:
        /// 1. leaf nodes, and
        /// 2. strictly data-containing (no formulas).
        /// </summary>
        /// <returns></returns>
        public TreeNode[] TerminalInputCells()
        {
            // this has to be done via recursive descent; a simple LINQ expression
            // will not suffice.
            throw new NotImplementedException();
            return TerminalInputNodes().SelectMany(pair => pair.getInputs())
                                       .Where(node => !node.isFormula() && node.isLeaf())
                                       .ToArray();
        }

        public string ToDOT()
        {
            var visited = new HashSet<AST.Address>();
            String s = "digraph spreadsheet {\n";
            foreach (KeyValuePair<AST.Address,TreeNode> pair in formula_nodes)
            {
                s += pair.Value.ToDOT(visited);
            }
            return s + "\n}";
        }

        public bool ContainsLoop()
        {
            var OK = true;
            foreach (KeyValuePair<AST.Address, TreeNode> pair in formula_nodes)
            {
                // a loop is when we see the same node twice while recursing
                var visited_from = new Dictionary<TreeNode,TreeNode>();
                OK = OK && !pair.Value.ContainsLoop(visited_from, null);
            }
            return !OK;
        }
    }
}
