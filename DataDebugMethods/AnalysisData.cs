using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
using TreeDict = System.Collections.Generic.Dictionary<AST.Address, DataDebugMethods.TreeNode>;
using System.Diagnostics;

namespace DataDebugMethods
{
    public class AnalysisData
    {
        public List<TreeNode> nodelist;     // holds all the TreeNodes in the Excel fileTreeNode
        public TreeDict input_ranges;
        private bool no_progress;
        private ProgBar pb;
        public TreeDict formula_nodes;
        public TreeDict cell_nodes;
        public Excel.Sheets charts;

        public const int PROGRESS_LOW = 0;
        public const int PROGRESS_HIGH = 100;

        private int _pb_max;
        private int _pb_count = 0;

        public AnalysisData(Excel.Application application, Excel.Workbook wb, bool dont_show_progbar)
        {
            no_progress = dont_show_progbar;
            charts = wb.Charts;
            nodelist = new List<TreeNode>();
            input_ranges = new TreeDict();
            cell_nodes = new TreeDict();

            // Create a progress bar
            if (!no_progress)
            {
                pb = new ProgBar(PROGRESS_LOW, PROGRESS_HIGH);
                pb.SetProgress(0);
            }
        }

        public void SetProgress(int i)
        {
            Debug.Assert(i >= PROGRESS_LOW && i <= PROGRESS_HIGH);
            if (!no_progress) pb.SetProgress(i);
        }

        public void SetPBMax(int max)
        {
            _pb_max = max;
        }

        public void PokePB()
        {
            _pb_count += 1;
            this.SetProgress(_pb_count * 100 / _pb_max);
        }

        public void KillPB()
        {
            // Kill progress bar
            if (!no_progress) pb.Close();
        }

        public TreeNode[] TerminalFormulaNodes()
        {
            // return only the formula nodes which do not provide
            // input to any other cell and which are also not
            // in our list of excluded functions
            return formula_nodes.Where(pair => pair.Value.getChildren().Count == 0)
                                .Select(pair => pair.Value).ToArray();
        }

        public TreeNode[] TerminalInputNodes()
        {
            // this should filter out the following two cases:
            // 1. input range is intermediate (acts as input to a formula
            //    and also contains a formula which consumes input from
            //    another range).
            // 2. the range is actually a formula cell
            return input_ranges.Where(pair => !pair.Value.GetDontPerturb())
                               .Select(pair => pair.Value).ToArray();
        }
    }
}
