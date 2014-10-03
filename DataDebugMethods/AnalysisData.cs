//using System;
//using System.Collections.Generic;
//using System.Linq;
//using System.Text;
//using Excel = Microsoft.Office.Interop.Excel;
//using CellDict = System.Collections.Generic.Dictionary<AST.Address, AST.COMRef>;
//using RangeDict = System.Collections.Generic.Dictionary<AST.Range, AST.COMRef>;
//using System.Diagnostics;

//namespace DataDebugMethods
//{
//    public class AnalysisData
//    {
//        public List<AST.COMRef> nodelist;
//        public RangeDict input_ranges;
//        public CellDict formula_nodes;
//        public CellDict cell_nodes;
//        public double tree_construct_time;
//        private ProgBar pb;

//        private int _pb_max;
//        private int _pb_count = 0;

//        public AnalysisData(Excel.Application application, Excel.Workbook wb) 
//        {
//            nodelist = new List<AST.COMRef>();
//            cell_nodes = new CellDict();
//        }

//        public AnalysisData(Excel.Application application, Excel.Workbook wb, ProgBar progbar)
//            : this (application, wb)
//        {
//            pb = progbar;
//        }

//        public void SetProgress(int i)
//        {
//            if (pb != null) pb.SetProgress(i);
//        }

//        public void SetPBMax(int max)
//        {
//            _pb_max = max;
//        }

//        public void PokePB()
//        {
//            if (pb != null)
//            {
//                _pb_count += 1;
//                this.SetProgress(_pb_count * 100 / _pb_max);
//            }
//        }

//        private void KillPB()
//        {
//            // Kill progress bar
//            if (pb != null) pb.Close();
//        }

//        public AST.COMRef[] TerminalFormulaNodes(bool all_outputs)
//        {
//            // return only the formula nodes which do not provide
//            // input to any other cell and which are also not
//            // in our list of excluded functions
//            if (all_outputs)
//            {
//                return formula_nodes.Select(pair => pair.Value).ToArray();
//            }
//            else
//            {
//                return formula_nodes.Where(pair => pair.Value.getOutputs().Count == 0)
//                                    .Select(pair => pair.Value).ToArray();
//            }
//        }

//        public COMRef[] TerminalInputNodes()
//        {
//            // this should filter out the following two cases:
//            // 1. input range is intermediate (acts as input to a formula
//            //    and also contains a formula which consumes input from
//            //    another range).
//            // 2. the range is actually a formula cell
//            return input_ranges.Where(pair => !pair.Value.DoNotPerturb
//                                              && !pair.Value.IsFormula)
//                               .Select(pair => pair.Value).ToArray();
//        }

//        /// <summary>
//        /// This method returns all input TreeNodes that are guaranteed to be:
//        /// 1. leaf nodes, and
//        /// 2. strictly data-containing (no formulas).
//        /// </summary>
//        /// <returns>TreeNode[]</returns>
//        public COMRef[] TerminalInputCells()
//        {
//            // this folds all of the inputs for all of the
//            // outputs into a set of distinct data-containing cells
//            var iecells = TerminalFormulaNodes(true).Aggregate(
//                            Enumerable.Empty<COMRef>(),
//                            (acc, node) => acc.Union<COMRef>(getChildCells(node))
//                          );
//            return iecells.ToArray<COMRef>();
//        }

//        /// <summary>
//        ///  This method returns all TreeNodes cells that participate in a computation.  Note
//        ///  that these nodes may be formulas!
//        /// </summary>
//        /// <returnsTreeNode[]></returns>
//        public COMRef[] allComputationCells()
//        {
//            // this folds all of the inputs for all of the
//            // outputs into a set of distinct data-containing cells
//            var iecells = TerminalFormulaNodes(true).Aggregate(
//                            Enumerable.Empty<COMRef>(),
//                            (acc, node) => acc.Union<COMRef>(getAllCells(node))
//                          );
//            return iecells.ToArray<COMRef>();
//        }

//        private IEnumerable<COMRef> getAllCells(COMRef node)
//        {
//            var thiscell = node;
//            var children = node.getInputs().SelectMany(n => getAllCells(n));
//            List<COMRef> results = new List<COMRef>(children);
//            results.Add(thiscell);
//            return results;
//        }

//        private IEnumerable<COMRef> getChildCells(COMRef node)
//        {
//            // base case: node is a cell (not a range), it has no children, and it's not a formula
//            if (node.IsCell && node.getInputs().Count() == 0 && !node.IsFormula) {
//                return new List<COMRef> { node };
//            } else {
//            // recursive case: node *may* have children; if so, recurse
//                var children = node.getInputs().SelectMany(n => getChildCells(n));
//                return children;
//            }
//        }
//    }
//}
