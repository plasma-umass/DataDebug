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
        //public List<TreeNode> originalColorNodes = new List<TreeNode>(); //List for storing the original colors for all nodes
        public List<TreeNode> nodelist;        //This is a list holding all the TreeNodes in the Excel file
        public double[][][][] impacts_grid; //This is a multi-dimensional array of doubles that will hold each cell's impact on each of the outputs
        public bool[][][][] reachable_grid; //This is a multi-dimensional array of bools that will indicate whether a certain output is reachable from a certain cell
        public double[][] min_max_delta_outputs; //This keeps the min and max delta for each output; first index indicates the output index; second index 0 is the min delta, 1 is the max delta for that output
        public List<TreeNode> input_ranges;      // This is a list of input ranges, with each Excel.Range COM object encapsulated in a TreeNode
        public List<StartValue> starting_outputs; //This will store the values of all the output nodes at the start of the procedure for swapping values (fuzzing)
        public List<TreeNode> output_cells; //This will store all the output nodes at the start of the fuzzing procedure
        public List<double[]>[] reachable_impacts_grid;  //This will store impacts for cells reachable from a particular output
        public double[][][] reachable_impacts_grid_array; //This will store impacts for cells reachable from a particular output in array form
        public double[][][] influences_grid;
        public int input_cells_in_computation_count = 0;
        public int raw_input_cells_in_computation_count = 0;
        public int formula_cells_count;
        public System.Diagnostics.Stopwatch global_stopwatch = new System.Diagnostics.Stopwatch();
        public ProgBar pb;
        public TreeDict formula_nodes;
        public TreeDict cell_nodes;
        public TimeSpan tree_building_timespan;
        public TimeSpan impact_scoring_timespan;
        public TimeSpan swapping_timespan;
        public int outliers_count; //This gets assigned and updated in the Analysis class
        public int[][][] times_perturbed;
        public Excel.Sheets worksheets;
        public Excel.Sheets charts;

        public const int PROGRESS_LOW = 0;
        public const int PROGRESS_HIGH = 100;

        private int _pb_max;
        private int _pb_count = 0;

        public AnalysisData(Excel.Application application)
        {
            worksheets = application.Worksheets;
            charts = application.Charts;
            nodelist = new List<TreeNode>();            // holds all the TreeNodes in the Excel file
            input_ranges = new List<TreeNode>();              // holds all the input ranges of TreeNodes in the Excel file
            starting_outputs = new List<StartValue>();  // holds the values of all the output nodes at the start of the procedure for swapping values (fuzzing)
            output_cells = new List<TreeNode>();        // holds the output nodes at the start of the fuzzing procedure
            cell_nodes = new TreeDict();
        }

        public void Reset()
        {
            // reset lists
            nodelist = new List<TreeNode>();
            input_ranges = new List<TreeNode>();
            starting_outputs = new List<StartValue>();
            output_cells = new List<TreeNode>();
            cell_nodes = new TreeDict();

            // Create a progress bar
            pb = new ProgBar(PROGRESS_LOW, PROGRESS_HIGH);
        }

        public void SetProgress(int i)
        {
            Debug.Assert(i >= PROGRESS_LOW && i <= PROGRESS_HIGH);
            pb.SetProgress(i);
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
            pb.Close();
        }
    }
}
