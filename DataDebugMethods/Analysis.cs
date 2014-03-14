using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
using TreeDict = System.Collections.Generic.Dictionary<AST.Address, DataDebugMethods.TreeNode>;
using TreeDictPair = System.Collections.Generic.KeyValuePair<AST.Address, DataDebugMethods.TreeNode>;
using TreeScore = System.Collections.Generic.Dictionary<DataDebugMethods.TreeNode, int>;
using Range = Microsoft.Office.Interop.Excel.Range;
using System.Diagnostics;
using Stopwatch = System.Diagnostics.Stopwatch;
using Microsoft.FSharp.Core;
using ExtensionMethods;

namespace DataDebugMethods
{
    public class Analysis
    {
        private static Dictionary<TreeNode, InputSample> StoreInputs(TreeNode[] inputs)
        {
            var d = new Dictionary<TreeNode, InputSample>();
            foreach (TreeNode input_range in inputs)
            {
                var com = input_range.getCOMObject();
                var s = new InputSample(input_range.Rows(), input_range.Columns());

                // store the entire COM array as a multiarray
                // in one fell swoop.
                s.AddArray(com.Value2);

                // add stored input to dict
                d.Add(input_range, s);

                // this is to force excel to recalculate its outputs
                // exactly the same way that it will for our bootstraps
                BootMemo.ReplaceExcelRange(com, s);
            }
            return d;
        }

        private static Dictionary<TreeNode, string> StoreOutputs(TreeNode[] outputs)
        {
            var d = new Dictionary<TreeNode, string>();
            foreach (TreeNode output_fn in outputs)
            {
                // we want to save the actual value of the function
                // since we don't know whether the function is string or numeric
                // until later, leave it as string for now
                d.Add(output_fn, output_fn.getCOMValueAsString());
            }
            return d;
        }

        public static InputSample[] Resample(int num_bootstraps, InputSample orig_vals, Random rng)
        {
            // the resampled values go here
            var ss = new InputSample[num_bootstraps];
            
            // sample with replacement to get i bootstrapped samples
            for (var i = 0; i < num_bootstraps; i++)
            {
                var s = new InputSample(orig_vals.Rows(), orig_vals.Columns());

                // make a vector of index counters
                var inc_count = new int[orig_vals.Length()];

                // randomly sample j values, with replacement
                for (var j = 0; j < orig_vals.Length(); j++)
                {
                    // randomly select a value from the original values
                    int input_idx = rng.Next(0, orig_vals.Length());
                    inc_count[input_idx] += 1;
                    Debug.Assert(input_idx < orig_vals.Length());
                    string value = orig_vals.GetInput(input_idx);
                    s.Add(value);
                }

                // indicate which indices are excluded
                s.SetIncludes(inc_count);

                // add the new InputSample to the output array
                ss[i] = s;
            }

            return ss;
        }

        // num_bootstraps: the number of bootstrap samples to get
        // inputs: a list of inputs; each TreeNode represents an entire input range
        // outputs: a list of outputs; each TreeNode represents a function
        public static TreeScore Bootstrap(int num_bootstraps, AnalysisData data, Excel.Application app, bool weighted)
        {
            // this modifies the weights of each node
            PropagateWeights(data);

            // filter out non-terminal functions
            var output_fns = data.TerminalFormulaNodes();
            // filter out non-terminal inputs
            var input_rngs = data.TerminalInputNodes();

            // first idx: the index of the TreeNode in the "inputs" array
            // second idx: the ith bootstrap
            var resamples = new InputSample[input_rngs.Length][];

            // RNG for sampling
            var rng = new Random();

            // we save initial inputs here
            var initial_inputs = StoreInputs(input_rngs);
            var initial_outputs = StoreOutputs(output_fns);

            // populate bootstrap array
            // for each input range (a TreeNode)
            System.Threading.Tasks.Parallel.For(0, input_rngs.Length, i =>
            {
                // this TreeNode
                var t = input_rngs[i];

                // resample
                resamples[i] = Resample(num_bootstraps, initial_inputs[t], rng);
            });

            // first idx: the fth function output
            // second idx: the ith input
            // third idx: the bth bootstrap
            var boots = ComputeBootstraps(num_bootstraps, initial_inputs, resamples, input_rngs, output_fns, data);

            // restore formulas
            foreach (TreeDictPair pair in data.formula_nodes)
            {
                TreeNode node = pair.Value;
                if (node.isFormula())
                {
                    node.getCOMObject().Formula = node.getFormula();
                }
            }

            // do appropriate hypothesis test, and add weighted test scores, and return result dict
            return ScoreInputs(input_rngs, output_fns, initial_outputs, boots, weighted);
        }

        public static TreeScore ScoreInputs(TreeNode[] input_rngs, TreeNode[] output_fns, Dictionary<TreeNode,string> initial_outputs, FunctionOutput<string>[][][] boots, bool weighted)
        {
            // dict of exclusion scores for each input CELL TreeNode
            var iexc_scores = new TreeScore();

            // convert bootstraps to numeric, if possible, sort in ascending order
            // then compute quantiles and test whether an input is an outlier
            // i is the index of the range in the input array; an ARRAY of CELLS
            for (int i = 0; i < input_rngs.Length; i++)
            {
                // f is the index of the function in the output array; a SINGLE CELL
                for (int f = 0; f < output_fns.Length; f++)
                {
                    // this function output treenode
                    TreeNode functionNode = output_fns[f];

                    // this function's input range treenode
                    TreeNode rangeNode = input_rngs[i];

                    // do the hypothesis test and then merge
                    // the scores from previous tests
                    TreeScore s;
                    if (FunctionOutputsAreNumeric(boots[f][i]))
                    {
                        s = NumericHypothesisTest(rangeNode, functionNode, boots[f][i], initial_outputs[functionNode], weighted);
                    }
                    else
                    {
                        s = StringHypothesisTest(rangeNode, functionNode, boots[f][i], initial_outputs[functionNode], weighted);
                    }
                    iexc_scores = DictAdd(iexc_scores, s);
                }
            }
            return iexc_scores;
        }

        public static TreeScore DictAdd(TreeScore d1, TreeScore d2)
        {
            var d3 = new TreeScore();
            if (d1 != null)
            {
                foreach (KeyValuePair<TreeNode, int> pair in d1)
                {
                    d3.Add(pair.Key, pair.Value);
                }
            }
            if (d2 != null)
            {
                foreach (KeyValuePair<TreeNode, int> pair in d2)
                {
                    int score;
                    if (d3.TryGetValue(pair.Key, out score))
                    {
                        d3[pair.Key] = score + pair.Value;
                    }
                    else
                    {
                        d3.Add(pair.Key, pair.Value);
                    }
                }
            }

            return d3;
        }

        public static TreeScore StringHypothesisTest(TreeNode rangeNode, TreeNode functionNode, FunctionOutput<string>[] boots, string initial_output, bool weighted)
        {
            // this function's input cells
            var input_cells = rangeNode.getInputs().ToArray();

            // scores
            var iexc_scores = new TreeScore();

            // exclude each index, in turn
            for (int i = 0; i < input_cells.Length; i++)
            {
                // default weight
                int weight = 1;

                // add weight to score if test fails
                TreeNode xtree = input_cells[i];
                if (weighted)
                {
                    // the weight of the function value of interest
                    weight = (int)functionNode.getWeight();
                }

                if (RejectNullHypothesis(boots, initial_output, i))
                {

                    if (iexc_scores.ContainsKey(xtree))
                    {
                        iexc_scores[xtree] += weight;
                    }
                    else
                    {
                        iexc_scores.Add(xtree, weight);
                    }
                }
                else
                {
                    // we need to at least add the value to the tree
                    if (!iexc_scores.ContainsKey(xtree))
                    {
                        iexc_scores.Add(xtree, 0);
                    }
                }
            }

            return iexc_scores;
        }

        public static TreeScore NumericHypothesisTest(TreeNode rangeNode, TreeNode functionNode, FunctionOutput<string>[] boots, string initial_output, bool weighted)
        {
            // this function's input cells
            var input_cells = rangeNode.getInputs().ToArray();

            // scores
            var input_exclusion_scores = new TreeScore();

            // convert to numeric
            var numeric_boots = ConvertToNumericOutput(boots);

            // sort
            var sorted_num_boots = SortBootstraps(numeric_boots);

            // for each excluded index, test whether the original input
            // falls outside our bootstrap confidence bounds
            for (int i = 0; i < input_cells.Length; i++)
            {
                // default weight
                int weight = 1;

                // add weight to score if test fails
                TreeNode xtree = input_cells[i];
                if (weighted)
                {
                    // the weight of the function value of interest
                    weight = (int)functionNode.getWeight();
                }

                double outlieriness = RejectNullHypothesis(sorted_num_boots, initial_output, i);

                if (outlieriness != 0.0)
                {
                    // get the xth indexed input in input_rng i
                    if (input_exclusion_scores.ContainsKey(xtree))
                    {
                        input_exclusion_scores[xtree] += (int)(weight * outlieriness);
                    }
                    else
                    {
                        input_exclusion_scores.Add(xtree, (int)(weight * outlieriness));
                    }
                }
                else
                {
                    // we need to at least add the value to the tree
                    if (!input_exclusion_scores.ContainsKey(xtree))
                    {
                        input_exclusion_scores.Add(xtree, 0);
                    }
                }
            }
            return input_exclusion_scores;
        }

        public static AST.Address GetTopOutlier(IEnumerable<Tuple<double, TreeNode>> quantiles, HashSet<AST.Address> known_good, double significance)
        {
            if (quantiles.Count() == 0)
            {
                return null;
            }

            //only flag quantiles that begin past the significance cutoff
            //identify the quantile which straddles the significance cutoff
            double last_excluded_quantile = 1.0;
            foreach (var q in quantiles)
            {
                if (q.Item1 >= significance)
                {
                    last_excluded_quantile = q.Item1;
                    break;
                }
            }

            // filter out cells below our significance level
            var significant_scores = quantiles.Where(tup => tup.Item1 > last_excluded_quantile);

            // filter out cells marked as OK
            var filtered_scores = significant_scores.Where(tup => !known_good.Contains(tup.Item2.GetAddress()));

            if (filtered_scores.Count() != 0)
            {
                // get TreeNode corresponding to most unusual score
                var flagged_cell = filtered_scores.Last().Item2;

                // return cell address
                return flagged_cell.GetAddress();
            }
            else
            {
                return null;
            }
        }

        public static AST.Address FlagTopOutlier(IEnumerable<Tuple<double,TreeNode>> quantiles, HashSet<AST.Address> known_good, double significance, Excel.Application app)
        {
            var flagged_cell = GetTopOutlier(quantiles, known_good, significance);

            if (flagged_cell != null)
            {
                // get COM object for cell
                var comcell = flagged_cell.GetCOMObject(app);

                // highlight cell
                comcell.Interior.Color = System.Drawing.Color.Red;
            }

            // return cell address
            return flagged_cell;
        }

        public static void ColorOutliers(TreeScore input_exclusion_scores)
        {
            var f_input_scores = input_exclusion_scores;

            // find value of the max element; we use this to calibrate our scale
            //double min_score = input_exclusion_scores.Select(pair => pair.Value).Min();  // min value is always zero
            double max_score = f_input_scores.Select(pair => pair.Value).Max();  // largest value we've seen

            // if the user is using the tool in iterative mode, only highlight the
            // highest-ranked cell that is not in the known_good cells

            // highlight all of them
            double min_score = max_score;   //smallest value we've seen
            foreach (KeyValuePair<TreeNode, int> pair in input_exclusion_scores)
            {
                if (pair.Value < min_score && pair.Value != 0)
                {
                    min_score = pair.Value;
                }
            }
            if (min_score == max_score)
            {
                min_score = 0;
            }

            min_score = 0.50 * min_score; //this is so that the smallest outlier also gets colored, rather than being white

            // calculate the color of each cell
            string outlierValues = "";
            foreach (KeyValuePair<TreeNode, int> pair in f_input_scores)
            {
                var cell = pair.Key;

                int cval = 0;
                // this happens when there are no suspect inputs.
                if (max_score - min_score == 0)
                {
                    cval = 0;
                }
                else
                {
                    if (pair.Value != 0)
                    {
                        //cval = (int)(255 * (Math.Pow(1.01, pair.Value) - Math.Pow(1.01, min_score)) / (Math.Pow(1.01, max_score) - Math.Pow(1.01, min_score)));
                        cval = (int)(255 * (pair.Value - min_score) / (max_score - min_score));
                        outlierValues += cell.getCOMObject().Address + " : " + pair.Value + ";\t" + cval + Environment.NewLine;
                    }
                }
                // to make something a shade of red, we set the "red" value to 255, and adjust the OTHER values.
                // if cval == 0, skip, because otherwise we end up coloring it white
                if (cval != 0)
                {
                    var color = System.Drawing.Color.FromArgb(255, 255, 255 - cval, 255 - cval);
                    cell.getCOMObject().Interior.Color = color;
                }
            }
        }

        // initializes the first and second dimensions
        private static FunctionOutput<string>[][][] InitJagged3DBootstrapArray(int fn_idx_sz, int o_idx_sz, int b_idx_sz)
        {
            var bs = new FunctionOutput<string>[fn_idx_sz][][];
            for (int f = 0; f < fn_idx_sz; f++)
            {
                bs[f] = new FunctionOutput<string>[o_idx_sz][];
                for (int o = 0; o < o_idx_sz; o++)
                {
                    bs[f][o] = new FunctionOutput<string>[b_idx_sz];
                }
            }
            return bs;
        }

        private static FunctionOutput<string>[][][] ComputeBootstraps(int num_bootstraps, Dictionary<TreeNode, InputSample> initial_inputs, InputSample[][] resamples,
                                                                      TreeNode[] input_arr, TreeNode[] output_arr, AnalysisData data)
        {
            // first idx: the fth function output
            // second idx: the ith input
            // third idx: the bth bootstrap
            var bootstraps = InitJagged3DBootstrapArray(output_arr.Length, input_arr.Length, num_bootstraps);

            // Set progress bar max
            int maxcount = num_bootstraps * input_arr.Length;
            data.SetPBMax(maxcount);

            // init bootstrap memo
            var bootsaver = new BootMemo[input_arr.Length];

            // DEBUG
            var hits = 0;
            var sw = new Stopwatch();
            sw.Start();

            // compute function outputs for each bootstrap
            // inputs[i] is the ith input range
            for (var i = 0; i < input_arr.Length; i++)
            {
                var t = input_arr[i];
                var com = t.getCOMObject();
                bootsaver[i] = new BootMemo();
                            
                // replace the values of the COM object with the jth bootstrap,
                // save all function outputs, and
                // restore the original input
                for (var b = 0; b < num_bootstraps; b++)
                {
                    // use memo DB
                    FunctionOutput<string>[] fos = bootsaver[i].FastReplace(com, initial_inputs[t], resamples[i][b], output_arr, ref hits, false);
                    for (var f = 0; f < output_arr.Length; f++)
                    {
                        bootstraps[f][i][b] = fos[f];
                    }
                    data.PokePB();
                }

                // restore the COM value; faster to do once, at the end (this saves n-1 replacements)
                BootMemo.ReplaceExcelRange(com, initial_inputs[t]);
            }

            // Kill progress bar
            data.KillPB();

            sw.Stop();

            return bootstraps;
        }

        // are all of the values numeric?
        public static bool FunctionOutputsAreNumeric(FunctionOutput<string>[] boots)
        {
            for (int b = 0; b < boots.Length; b++)
            {
                if (!ExcelParser.isNumeric(boots[b].GetValue()))
                {
                    return false;
                }
            }
            return true;
        }

        // attempts to convert all of the bootstraps for FunctionOutput[function_idx, input_idx, _] to doubles
        public static FunctionOutput<double>[] ConvertToNumericOutput(FunctionOutput<string>[] boots)
        {
            var fi_boots = new FunctionOutput<double>[boots.Length];

            for (int b = 0; b < boots.Length; b++)
            {
                FunctionOutput<string> boot = boots[b];
                double value = System.Convert.ToDouble(boot.GetValue());
                fi_boots[b] = new FunctionOutput<double>(value, boot.GetExcludes());
            }
            return fi_boots;
        }

        // Sort numeric bootstrap values
        public static FunctionOutput<double>[] SortBootstraps(FunctionOutput<double>[] boots)
        {
            return boots.OrderBy(b => b.GetValue()).ToArray();
        }

        // Count instances of unique string output values and return multinomial probability vector
        public static Dictionary<string, double> BootstrapFrequency(FunctionOutput<string>[] boots)
        {
            var counts = new Dictionary<string, int>();

            foreach (FunctionOutput<string> boot in boots)
            {
                string key = boot.GetValue();
                int count;
                if (counts.TryGetValue(key, out count))
                {
                    counts[key] = count + 1;
                }
                else
                {
                    counts.Add(key, 1);
                }
            }

            var p_values = new Dictionary<string,double>();

            foreach (KeyValuePair<string,int> pair in counts)
            {
                p_values.Add(pair.Key, (double)pair.Value / (double)boots.Length);
            }

            return p_values;
        }

        // Exclude specified input index, compute multinomial probabilty vector, and return true if probability is below threshold
        public static bool RejectNullHypothesis(FunctionOutput<string>[] boots, string original_output, int exclude_index)
        {
            // filter bootstraps which include exclude_index
            var boots_exc = boots.Where(b => b.GetExcludes().Contains(exclude_index)).ToArray();

            // get p_value vector
            var freq = BootstrapFrequency(boots_exc);

            // what is the probability of seeing the original output?
            double p_val;
            if (!freq.TryGetValue(original_output, out p_val))
            {
                p_val = 0.0;
            }

            // test H_0
            return p_val < 0.05;
        }

        // Exclude a specified input index, compute quantiles, and check position of original input
        public static double RejectNullHypothesis(FunctionOutput<double>[] boots, string original_output, int exclude_index)
        {
            // filter bootstraps which include exclude_index
            var boots_exc = boots.Where(b => b.GetExcludes().Contains(exclude_index)).ToArray();
            //return neutral (0.5) if we are having a sparsity problem
            if (boots_exc.Length == 0)
            {
                return 0.5;
            }
            // index for value greater than 2.5% of the lowest values; we want to round down here
            var low_index = System.Convert.ToInt32(Math.Floor((float)(boots_exc.Length - 1) * .025));
            // index for value greater than 97.5% of the lowest values; we want to round up here
            var high_index = System.Convert.ToInt32(Math.Ceiling((float)(boots_exc.Length - 1) * 0.975));

            var low_value = boots_exc[low_index].GetValue();
            var high_value = boots_exc[high_index].GetValue();

            var original_output_d = System.Convert.ToDouble(original_output);

            // truncate the values to deal with floating point imprecision
            var low_value_tr = Math.Truncate(low_value * 10000) / 10000;
            var high_value_tr = Math.Truncate(high_value * 10000) / 10000;
            var original_tr = Math.Truncate(original_output_d * 10000) / 10000;
            
            // reject or fail to reject H_0
            if (high_value_tr != low_value_tr)
            {
                if (original_tr < low_value_tr)
                {
                    return Math.Abs((original_tr - low_value_tr) / Math.Abs(high_value_tr - low_value_tr)) * 100.0;
                }
                else if (original_tr > high_value_tr)
                {
                    return Math.Abs((original_tr - high_value_tr) / Math.Abs(high_value_tr - low_value_tr)) * 100.0;
                }
            }
            else if (high_value_tr != original_tr || low_value_tr != original_tr)
            {
                if (original_tr < low_value_tr)
                {
                    return Math.Abs(original_tr - low_value_tr) * 100.0;
                }
                else if (original_tr > high_value_tr)
                {
                    return Math.Abs(original_tr - high_value_tr) * 100.0;
                }
            }
            return 0.0;
        }

        // Computes quantile array.  Accepts key,value pairs so that arbitrary data can be kept and passed along.
        // Note that the Tuple's key type (K) is the basis for the quantile computation.
        public static IEnumerable<Tuple<double,V>> ComputeQuantile<K,V>(IEnumerable<Tuple<K,V>> inputs)
        {
            // sort values
            var sorted_values = inputs.OrderBy(tup => tup.Item1).ToArray();

            // init output list
            var outputs = new List<Tuple<double, V>>();

            // in loop, choose the next value, look for repeats of that value,
            // increment your pointer to the last instance of the value,
            // and then calculate the proportion of values to the left of the pointer (inclusive)
            int index = 0;
            while (index < sorted_values.Length)
            {
                // get current value
                var current_value = sorted_values[index].Item1;

                while (index + 1 < sorted_values.Length && current_value.Equals(sorted_values[index + 1].Item1))
                {
                    index += 1;
                }

                // calculate proportion of values to the left of the ptr
                var quantile = (double)(index + 1) / (double)sorted_values.Length;

                // update output with value
                outputs.Add(new Tuple<double,V>(quantile, sorted_values[index].Item2));

                index += 1;
            }

            return outputs;
        }

        // Propagate weights
        private static void PropagateWeights(AnalysisData data)
        {
            // starting set of functions; roots in the forest
            var functions = data.TerminalFormulaNodes();

            // for each forest
            foreach (TreeNode fn in functions)
            {
                fn.setWeight(PropagateTreeNodeWeight(fn));
            }
        }

        private static int PropagateTreeNodeWeight(TreeNode t)
        {
            var inputs = t.getInputs();
            // if we have no inputs, then we ARE an input
            if (inputs.Count() == 0)
            {
                t.setWeight(1);
                return 1;
            }
            // otherwise we have inputs, recursively compute their weights
            // and add to this one
            else
            {
                var weight = 0;
                foreach (var input in inputs)
                {
                    weight += PropagateTreeNodeWeight(input);
                }
                t.setWeight(weight);
                return weight;
            }
        }
    }
}
