using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
using TreeDict = System.Collections.Generic.Dictionary<AST.Address, DataDebugMethods.TreeNode>;
using TreeDictPair = System.Collections.Generic.KeyValuePair<AST.Address, DataDebugMethods.TreeNode>;
using Range = Microsoft.Office.Interop.Excel.Range;
using System.Diagnostics;

namespace DataDebugMethods
{
    public class Analysis
    {
        public static void perturbationAnalysis(ProgBar pb, 
                                            ref int[][][] times_perturbed, 
                                            Excel.Sheets worksheets, 
                                            Excel.Sheets charts, 
                                            int outliers_count, 
                                            List<TreeNode> ranges, 
                                            ref double[][] min_max_delta_outputs, 
                                            List<TreeNode> output_cells, 
                                            ref double[][][][] impacts_grid, 
                                            ref bool[][][][] reachable_grid,
                                            ref List<double[]>[] reachable_impacts_grid, 
                                            ref double[][][] reachable_impacts_grid_array, 
                                            TreeDict nodes,
                                            ref int raw_input_cells_in_computation_count, 
                                            List<StartValue> starting_outputs, 
                                            ref int input_cells_in_computation_count)
        {
            pb.SetProgress(25);

            //Grids for storing influences
            double[][][] influences_grid = null;
            times_perturbed = null;
            //influences_grid and times_perturbed are passed by reference so that they can be modified in the setUpGrids method
            ConstructTree.setUpGrids(ref influences_grid, ref times_perturbed, worksheets, charts);

            outliers_count = 0;
            //Procedure for swapping values within ranges, one cell at a time
            //if (!checkBox2.Checked) //Checks if the option for swapping values simultaneously is checked (not checked by default)
            //{
            List<TreeNode> swap_domain;
            swap_domain = ranges;

            //Initialize min_max_delta_outputs
            min_max_delta_outputs = new double[output_cells.Count][];
            for (int i = 0; i < output_cells.Count; i++)
            {
                min_max_delta_outputs[i] = new double[2];
                min_max_delta_outputs[i][0] = -1.0;
                min_max_delta_outputs[i][1] = 0.0;
            }

            //Initialize impacts_grid 
            //Initialize reachable_grid
            impacts_grid = new double[worksheets.Count][][][];
            reachable_grid = new bool[worksheets.Count][][][];
            foreach (Excel.Worksheet worksheet in worksheets)
            {
                impacts_grid[worksheet.Index - 1] = new double[worksheet.UsedRange.Rows.Count + worksheet.UsedRange.Row][][];
                reachable_grid[worksheet.Index - 1] = new bool[worksheet.UsedRange.Rows.Count + worksheet.UsedRange.Row][][];
                for (int row = 0; row < (worksheet.UsedRange.Rows.Count + worksheet.UsedRange.Row); row++)
                {
                    impacts_grid[worksheet.Index - 1][row] = new double[worksheet.UsedRange.Columns.Count + worksheet.UsedRange.Column][];
                    reachable_grid[worksheet.Index - 1][row] = new bool[worksheet.UsedRange.Columns.Count + worksheet.UsedRange.Column][];
                    for (int col = 0; col < (worksheet.UsedRange.Columns.Count + worksheet.UsedRange.Column); col++)
                    {
                        impacts_grid[worksheet.Index - 1][row][col] = new double[output_cells.Count];
                        reachable_grid[worksheet.Index - 1][row][col] = new bool[output_cells.Count];
                        for (int i = 0; i < output_cells.Count; i++)
                        {
                            impacts_grid[worksheet.Index - 1][row][col][i] = 0.0;
                            reachable_grid[worksheet.Index - 1][row][col][i] = false;
                        }
                    }
                }
            }

            //Initialize reachable_impacts_grid
            reachable_impacts_grid = new List<double[]>[output_cells.Count];
            for (int i = 0; i < output_cells.Count; i++)
            {
                reachable_impacts_grid[i] = new List<double[]>();
            }

            //Propagate weights  -- find the weights of all outputs and set up the reachable_grid entries
            foreach (TreeDictPair tdp in nodes)
            {
                var node = tdp.Value;
                if (!node.hasParents())
                {
                    node.setWeight(1.0);  //Set the weight of all input nodes to 1.0 to start
                    //Now we propagate it's weight to all of it's children
                    TreeNode.propagateWeightUp(node, 1.0, node, output_cells, reachable_grid, reachable_impacts_grid);
                    raw_input_cells_in_computation_count++;
                }
            }

            //Convert reachable_impacts_grid to array form
            reachable_impacts_grid_array = new double[output_cells.Count][][];
            for (int i = 0; i < output_cells.Count; i++)
            {
                reachable_impacts_grid_array[i] = reachable_impacts_grid[i].ToArray();
            }
            pb.SetProgress(40);
            ConstructTree.SwappingProcedure(swap_domain, ref input_cells_in_computation_count, ref min_max_delta_outputs, ref impacts_grid, ref times_perturbed, ref output_cells, ref reachable_grid, ref starting_outputs, ref reachable_impacts_grid_array);
           
            //Stop timing swapping procedure:
            pb.SetProgress(80);
        } //perturbationAnalysis ends here

        public static void outlierAnalysis(ProgBar pb, 
                                            int[][][] times_perturbed, 
                                            Excel.Sheets worksheets, 
                                            int outliers_count, 
                                            List<TreeNode> output_cells, 
                                            double[][][][] impacts_grid, 
                                            double[][][] reachable_impacts_grid_array)
        {
            ConstructTree.ComputeZScoresAndFindOutliers(output_cells, reachable_impacts_grid_array, impacts_grid, times_perturbed, worksheets, outliers_count);
            //Stop timing the zscore computation and outlier finding
            pb.SetProgress(pb.maxProgress());
            pb.Close();
            
            // Format and display the TimeSpan value. 
            //string tree_building_time = tree_building_timespan.TotalSeconds + ""; //String.Format("{0:00}:{1:00}.{2:00}", tree_building_timespan.Minutes, tree_building_timespan.Seconds, tree_building_timespan.Milliseconds / 10);
            //string swapping_time = (swapping_timespan.TotalSeconds - tree_building_timespan.TotalSeconds) + ""; //String.Format("{0:00}:{1:00}.{2:00}", swapping_timespan.Minutes, swapping_timespan.Seconds, swapping_timespan.Milliseconds / 10);
            //string impact_scoring_time = (impact_scoring_timespan.TotalSeconds - swapping_timespan.TotalSeconds) + ""; //String.Format("{0:00}:{1:00}.{2:00}", z_score_timespan.Minutes, z_score_timespan.Seconds, z_score_timespan.Milliseconds / 10);
            //global_stopwatch.Stop();
            // Get the elapsed time as a TimeSpan value.
            //TimeSpan global_timespan = global_stopwatch.Elapsed;
            //string global_time = String.Format("{0:00}:{1:00}:{2:00}.{3:00}", global_timespan.Hours, global_timespan.Minutes, global_timespan.Seconds, global_timespan.Milliseconds / 10);
            //string global_time = global_timespan.TotalSeconds + ""; //(tree_building_timespan.TotalSeconds + swapping_timespan.TotalSeconds + z_score_timespan.TotalSeconds + average_z_score_timespan.TotalSeconds + outlier_detection_timespan.TotalSeconds + outlier_coloring_timespan.TotalSeconds) + ""; //String.Format("{0:00}:{1:00}.{2:00}",

            //Display timeDisplay = new Display();
            //stats_text += "" //+ "Benchmark:\tNumber of formulas:\tRaw input count:\tInputs to computations:\tTotal (s):\tTree Construction (s):\tSwapping (s):\tZ-Score Calculation (s):\t"
                //  + "Outlier Detection (s):\tOutlier Coloring (s):\t"
                //+ "Outliers found:\n"
                //"Formula cells:\t" + formula_cells_count + "\n"
                //+ "Number of input cells involved in computations:\t" + input_cells_in_computation_count
                //+ "\nExecution times (seconds): "
                //+ Globals.ThisAddIn.Application.ActiveWorkbook.Name + "\t"
                //+ formula_cells_count + "\t"
                //+ raw_input_cells_in_computation_count + "\t"
                //+ input_cells_in_computation_count + "\t"
                //+ global_time + "\t"
                //+ tree_building_time + "\t"
                //+ swapping_time + "\t"
                //+ impact_scoring_time + "\t"
                //+ outliers_count;
            //timeDisplay.textBox1.Text = stats_text;
            //timeDisplay.ShowDialog();

        } //outlierAnalysis ends here

        private static Dictionary<TreeNode, InputSample> StoreInputs(List<TreeNode> inputs)
        {
            var d = new Dictionary<TreeNode, InputSample>();
            foreach (TreeNode input_range in inputs)
            {
                // the values are stored in this range's "parents"
                // (i.e., the actual cells)
                List<TreeNode> cells = input_range.getParents();
                var s = new InputSample(cells.Count);

                // store each input cell's contens
                foreach (TreeNode c in cells)
                {
                    s.Add(System.Convert.ToString(c.getCOMObject().Value2));
                }
                // add stored input to dict
                d.Add(input_range, s);
            }
            return d;
        }

        public class FunctionOutput
        {
            private string _value;
            private HashSet<int> _excludes;
            public FunctionOutput(string value, HashSet<int> excludes)
            {
                _value = value;
                _excludes = excludes;
            }
        }

        public class InputSample
        {
            private int _i = 0;             // internal length counter
            private string[] _input_array;  // the actual values of this array
            private HashSet<int> _excludes;    // list of inputs excluded in this sample

            public InputSample(int size)
            {
                _input_array = new string[size];
                _excludes = new HashSet<int>(Enumerable.Range(0, size));
            }
            public void Add(string value)
            {
                Debug.Assert(_i < _input_array.Length);
                _input_array[_i] = value;
                _i++;
            }
            public string GetInput(int num)
            {
                Debug.Assert(num < _input_array.Length);
                return _input_array[num];
            }
            public int Length()
            {
                return _i;
            }
            public HashSet<int> GetExcludes()
            {
                return _excludes;
            }
            public void SetExcludes(HashSet<int> exc)
            {
                _excludes = exc;
            }
        }

        public static InputSample[] Resample(int num_bootstraps, InputSample orig_vals, Random rng)
        {
            // the resampled values go here
            var ss = new InputSample[num_bootstraps];

            // sample with replacement to get i
            // bootstrapped samples
            for (var i = 0; i < num_bootstraps; i++)
            {
                var s = new InputSample(orig_vals.Length());

                // make a list of possibly-excluded indices
                var exc = new HashSet<int>(Enumerable.Range(0, orig_vals.Length()));

                // randomly sample j values, with replacement
                for (var j = 0; j < orig_vals.Length(); j++)
                {
                    // randomly select a value from the original values
                    int input_idx = rng.Next(0, orig_vals.Length());
                    exc.Remove(input_idx);
                    string value = orig_vals.GetInput(input_idx);
                    s.Add(value);
                }

                // indicate which indices are excluded
                s.SetExcludes(exc);
            }

            return ss;
        }

        public static void ReplaceExcelRange(Range com, InputSample input)
        {
            var i = 0;
            foreach (Range cell in com)
            {
                cell.Value2 = input.GetInput(i);
                i++;
            }
        }

        private static FunctionOutput[,] ComputeBootstraps(int num_bootstraps,
                                                           List<TreeNode> inputs,
                                                           List<TreeNode> outputs,
                                                           Dictionary<TreeNode, InputSample> initial_inputs,
                                                           InputSample[][] resamples)
        {
            // first idx: the output range idx in "outputs"
            // second idx: the ith bootstrap
            var bootstraps = new FunctionOutput[outputs.Count, num_bootstraps];

            // compute function outputs for each bootstrap
            // each input
            for (var i = 0; i < inputs.Count; i++)
            {
                var t = inputs[i];
                var com = t.getCOMObject();

                // replace the values of the COM object with each
                // bootstrap and save all function outputs
                for (var j = 0; j < num_bootstraps; j++)
                {
                    // replace the COM value
                    ReplaceExcelRange(com, resamples[i][j]);

                    // grab all outputs
                    for (var k = 0; k < outputs.Count; k++)
                    {
                        // save the output
                        bootstraps[k, j] = new FunctionOutput(outputs[k].getCOMValueAsString(), resamples[i][j].GetExcludes());
                    }

                    // reset the COM value to its original state
                    ReplaceExcelRange(com, initial_inputs[t]);
                }
            }

            return bootstraps;
        }

        // num_bootstraps: the number of bootstrap samples to get
        // inputs: a list of inputs; each TreeNode represents an entire input range
        // outputs: a list of outputs; each TreeNode represents a function
        
        public static FunctionOutput[,] Bootstrap(int num_bootstraps, List<TreeNode> inputs, List<TreeNode> outputs)
        {
            // first idx: the index of the TreeNode in the "inputs" array
            // second idx: the ith bootstrap
            var resamples = new InputSample[inputs.Count][];

            // RNG for sampling
            var rng = new Random();

            // we save initial inputs here
            var initial_inputs = StoreInputs(inputs);

            // populate bootstrap array
            // for each input range (a TreeNode)
            for (var i = 0; i < inputs.Count; i++)
            {
                // this TreeNode
                var t = inputs[i];
                // resample
                resamples[i] = Resample(num_bootstraps, initial_inputs[t], rng);
            }

            // replace each input range with a resample and
            // gather all outputs
            return ComputeBootstraps(num_bootstraps, inputs, outputs, initial_inputs, resamples);
        }
    }
}
