using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Numerics;
using System.Threading;
using Excel = Microsoft.Office.Interop.Excel;
using TreeScore = System.Collections.Generic.Dictionary<AST.Address, int>;
using Range = Microsoft.Office.Interop.Excel.Range;
using System.Diagnostics;
using Stopwatch = System.Diagnostics.Stopwatch;
using Microsoft.FSharp.Core;
using ExtensionMethods;

namespace DataDebugMethods
{
    public class ContainsLoopException : Exception { }

    public class Analysis
    {
        private static Dictionary<AST.Range, InputSample> StoreInputs(AST.Range[] inputs, DAG dag)
        {
            var d = new Dictionary<AST.Range, InputSample>();
            foreach (AST.Range input_range in inputs)
            {
                var com = dag.getCOMRefForRange(input_range);
                var s = new InputSample(com.Height, com.Width);

                // store the entire COM array as a multiarray
                // in one fell swoop.
                s.AddArray(com.Range.Value2);

                // add stored input to dict
                d.Add(input_range, s);

                // this is to force excel to recalculate its outputs
                // exactly the same way that it will for our bootstraps
                BootMemo.ReplaceExcelRange(com.Range, s);
            }

            return d;
        }

        public static Dictionary<AST.Address, string> StoreOutputs(AST.Address[] outputs, DAG dag)
        {
            // output dict
            var d = new Dictionary<AST.Address, string>();

            // partition all of the output addresses by their worksheet
            var addr_groups = outputs.GroupBy(addr => dag.getCOMRefForAddress(addr).WorksheetName);

            // for each worksheet, do an array read of the formulas
            foreach (IEnumerable<AST.Address> ws_fns in addr_groups)
            {
                // get worksheet used range
                var fstcr = dag.getCOMRefForAddress(ws_fns.First());
                var rng = fstcr.Worksheet.UsedRange;

                // get used range dimensions
                var left = rng.Column;
                var right = rng.Columns.Count + left - 1;
                var top = rng.Row;
                var bottom = rng.Rows.Count + top - 1;

                // get names
                var wsname = new FSharpOption<string>(fstcr.WorksheetName);
                var wbname = new FSharpOption<string>(fstcr.WorkbookName);
                var path = fstcr.Path;

                // sometimes the used range is a range
                if (left != right || top != bottom)
                {
                    // y is the first index
                    // x is the second index
                    object[,] data = rng.Value2;    // fast array read

                    var x_del = left - 1;
                    var y_del = top - 1;

                    foreach (AST.Address addr in ws_fns)
                    {
                        // construct address in formulas array
                        var x = addr.X - x_del;
                        var y = addr.Y - y_del;

                        // get string
                        String s = System.Convert.ToString(data[y, x]);
                        if (String.IsNullOrWhiteSpace(s))
                        {
                            d.Add(addr, "");
                        }
                        else
                        {
                            d.Add(addr, s);
                        }
                    }
                }
                // and other times it is a single cell
                else
                {
                    // construct the appropriate AST.Address
                    AST.Address addr = AST.Address.NewFromR1C1(top, left, wsname, wbname, path);

                    // make certain that it is actually a string
                    String s = System.Convert.ToString(rng.Value2);

                    // add to dictionary, as appropriate
                    if (String.IsNullOrWhiteSpace(s))
                    {
                        d.Add(addr, "");
                    }
                    else
                    {
                        d.Add(addr, s);
                    }
                }
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
                    if (input_idx >= orig_vals.Length())
                    {
                        throw new Exception("input_idx >= orig_vals.Length()");
                    }
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
        public static TreeScore DataDebug(int num_bootstraps,
                                          DAG dag,
                                          Excel.Application app,
                                          bool weighted,
                                          bool all_outputs,
                                          long max_duration_in_ms,
                                          Stopwatch sw,
                                          double significance,
                                          ProgBar pb)
        {
            // this modifies the weights of each node
            PropagateWeights(dag);

            // filter out non-terminal functions
            var output_fns = dag.terminalFormulaNodes(all_outputs);
            // filter out non-terminal inputs
            var input_rngs = dag.terminalInputVectors();

            // first idx: the index of the TreeNode in the "inputs" array
            // second idx: the ith bootstrap
            var resamples = new InputSample[input_rngs.Length][];

            // RNG for sampling
            var rng = new Random();

            // we save initial inputs and outputs here
            var initial_inputs = StoreInputs(input_rngs, dag);
            var initial_outputs = StoreOutputs(output_fns, dag);

            // Set progress bar max
            pb.setMax(input_rngs.Length * 2);

            #region RESAMPLE

            // populate bootstrap array
            // for each input range (a TreeNode)
            for (int i = 0; i < input_rngs.Length; i++)
            {
                // this TreeNode
                var t = input_rngs[i];

                // resample
                resamples[i] = Resample(num_bootstraps, initial_inputs[t], rng);

                // update progress bar
                pb.IncrementProgress();
            }

            #endregion RESAMPLE

            #region INFERENCE
            return Inference(
                num_bootstraps,
                resamples,
                initial_inputs,
                initial_outputs,
                input_rngs,
                output_fns,
                dag,
                weighted,
                significance,
                pb);
            #endregion INFERENCE
        }

        public class DataDebugJob
        {
            private DAG _dag;
            private FunctionOutput<string>[][] _bs;
            private Dictionary<AST.Address, string> _initial_outputs;
            private AST.Range _input;
            private AST.Address[] _outputs;
            private bool _weighted;
            private double _significance;
            private ManualResetEvent _mre;
            private TreeScore _score; // dict of exclusion scores for each input CELL TreeNode

            public DataDebugJob(
                DAG dag,
                FunctionOutput<String>[][] bs,
                Dictionary<AST.Address, string> initial_outputs,
                AST.Range input,
                AST.Address[] output_arr,
                bool weighted,
                double significance,
                ManualResetEvent mre)
            {
                _dag = dag;
                _bs = bs;
                _initial_outputs = initial_outputs;
                _input = input;
                _outputs = output_arr;
                _weighted = weighted;
                _significance = significance;
                _mre = mre;
                _score = new TreeScore();
            }

            public TreeScore Result
            {
                get { return _score; }
            }

            private void hypothesisTests()
            {
                for (var f = 0; f < _outputs.Length; f++)
                {
                    AST.Address output = _outputs[f];

                    // do the hypothesis test and then merge
                    // the scores from previous tests
                    TreeScore s;
                    if (FunctionOutputsAreNumeric(_bs[f]))
                    {
                        s = NumericHypothesisTest(_dag, _input, output, _bs[f], _initial_outputs[output], _weighted, _significance);
                    }
                    else
                    {
                        s = StringHypothesisTest(_dag, _input, output, _bs[f], _initial_outputs[output], _weighted, _significance);
                    }
                    _score = DictAdd(_score, s);
                }
            }

            public void threadPoolCallback(Object threadContext)
            {
                // perform hypothesis tests
                hypothesisTests();

                // OK to dealloc fields; this object lives on because it is
                // needed for job control
                _bs = null;
                _initial_outputs = null;
                _input = null;
                _outputs = null;

                // notify
                _mre.Set();
            }
        }

        public static TreeScore Inference(
            int num_bootstraps,
            InputSample[][] resamples,
            Dictionary<AST.Range, InputSample> initial_inputs,
            Dictionary<AST.Address, string> initial_outputs,
            AST.Range[] input_arr,
            AST.Address[] output_arr,
            DAG dag,
            bool weighted,
            double significance,
            ProgBar pb)
        {
            // synchronization token
            object lock_token = new Object();

            // init thread event notification array
            var mres = new ManualResetEvent[input_arr.Length];

            // init job storage
            var ddjs = new DataDebugJob[input_arr.Length];

            // init started jobs count
            var sjobs = 0;

            // init completed jobs count
            var cjobs = 0;

            // last-ditch effort flag
            bool last_try = false;

            // init score storage
            var scores = new TreeScore();

            for (int i = 0; i < input_arr.Length; i++)
            {
                try
                {
                    #region BOOTSTRAP
                    // bootstrapping is done in the parent STA thread because
                    // the .NET threading model prohibits thread pools (which
                    // are MTA) from accessing STA COM objects directly.

                    // alloc bootstrap storage for each output (f), for each resample (b)
                    FunctionOutput<string>[][] bs = new FunctionOutput<string>[initial_outputs.Count][];
                    for (int f = 0; f < initial_outputs.Count; f++)
                    {
                        bs[f] = new FunctionOutput<string>[num_bootstraps];
                    }

                    // init memoization table for input vector i
                    var memo = new BootMemo();

                    // fetch the input range TreeNode
                    var input = input_arr[i];

                    // fetch the input range COM object
                    var com = dag.getCOMRefForRange(input).Range;

                    // compute outputs
                    // replace the values of the COM object with the jth bootstrap,
                    // save all function outputs, and
                    // restore the original input
                    for (var b = 0; b < num_bootstraps; b++)
                    {
                        // lookup outputs from memo table; otherwise do replacement, compute outputs, store them in table, and return them
                        FunctionOutput<string>[] fos = memo.FastReplace(com, dag, initial_inputs[input], resamples[i][b], output_arr, false);
                        for (var f = 0; f < output_arr.Length; f++)
                        {
                            bs[f][b] = fos[f];
                        }
                    }

                    // restore the original inputs; faster to do once, after bootstrapping is done
                    BootMemo.ReplaceExcelRange(com, initial_inputs[input]);

                    // TODO: restore formulas if it turns out that they were overwrittern
                    //       this should never be the case
                    #endregion BOOTSTRAP

                    #region HYPOTHESIS_TEST
                    // cancellation token
                    mres[i] = new ManualResetEvent(false);

                    // set up job
                    ddjs[i] = new DataDebugJob(
                                dag,
                                bs,
                                initial_outputs,
                                input_arr[i],
                                output_arr,
                                weighted,
                                significance,
                                mres[i]
                                );

                    sjobs++;

                    // hand job to thread pool
                    ThreadPool.QueueUserWorkItem(ddjs[i].threadPoolCallback, i);
                    #endregion HYPOTHESIS_TEST

                    // update progress bar
                    pb.IncrementProgress();
                }
                catch (System.OutOfMemoryException e)
                {
                    if (!last_try)
                    {
                        // If there are no more jobs running, but
                        // we still can't allocate memory, try invoking
                        // GC and then trying again
                        cjobs = mres.Count(mre => mre.WaitOne(0));
                        if (sjobs - cjobs == 0)
                        {
                            GC.Collect();
                            last_try = true;
                        }
                    }
                    else
                    {
                        // we just don't have enough memory
                        throw e;
                    }

                    // wait for any of the 0..i-1 work items
                    // to complete and try again
                    WaitHandle.WaitAny(mres.Take(i).ToArray());
                }
            }

            // Do not proceed until all hypothesis tests are done.
            // WaitHandle.WaitAll cannot be called on an STA thread which
            // is why we call WaitOne in a loop.
            // Merge scores as data becomes available.
            for (int i = 0; i < input_arr.Length; i++)
            {
                mres[i].WaitOne();
                scores = DictAdd(scores, ddjs[i].Result);
            }

            return scores;
        }

        public static TreeScore DictAdd(TreeScore d1, TreeScore d2)
        {
            var d3 = new TreeScore();
            if (d1 != null)
            {
                foreach (KeyValuePair<AST.Address, int> pair in d1)
                {
                    d3.Add(pair.Key, pair.Value);
                }
            }
            if (d2 != null)
            {
                foreach (KeyValuePair<AST.Address, int> pair in d2)
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

        public static TreeScore StringHypothesisTest(DAG dag, AST.Range rangeNode, AST.Address functionNode, FunctionOutput<string>[] boots, string initial_output, bool weighted, double significance)
        {
            // this function's input cells
            var input_cells = rangeNode.Addresses();

            // scores
            var iexc_scores = new TreeScore();

            var inputs_sz = input_cells.Count();

            // exclude each index, in turn
            for (int i = 0; i < inputs_sz; i++)
            {
                // default weight
                int weight = 1;

                // add weight to score if test fails
                AST.Address xtree = input_cells[i];
                if (weighted)
                {
                    // the weight of the function value of interest
                    weight = dag.getWeight(functionNode);
                }

                if (RejectNullHypothesis(boots, initial_output, i, significance))
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

        public static TreeScore NumericHypothesisTest(DAG dag, AST.Range rangeNode, AST.Address functionNode, FunctionOutput<string>[] boots, string initial_output, bool weighted, double significance)
        {
            // this function's input cells
            var input_cells = rangeNode.Addresses();

            var inputs_sz = input_cells.Count();

            // scores
            var input_exclusion_scores = new TreeScore();

            // convert to numeric
            var numeric_boots = ConvertToNumericOutput(boots);

            // sort
            var sorted_num_boots = SortBootstraps(numeric_boots);

            // for each excluded index, test whether the original input
            // falls outside our bootstrap confidence bounds
            for (int i = 0; i < inputs_sz; i++)
            {
                // default weight
                int weight = 1;

                // add weight to score if test fails
                AST.Address xtree = input_cells[i];
                if (weighted)
                {
                    // the weight of the function value of interest
                    weight = dag.getWeight(functionNode);
                }

                double outlieriness = RejectNullHypothesis(sorted_num_boots, initial_output, i, significance);

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

        // are all of the values numeric?
        public static bool FunctionOutputsAreNumeric(FunctionOutput<string>[] boots)
        {
            for (int i = 0; i < boots.Length; i++)
            {
                double d;
                if (!Double.TryParse(boots[i].GetValue(), out d))
                {
                    return false;
                }
            };
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
        public static Dictionary<string, double> BootstrapFrequency(IEnumerable<FunctionOutput<string>> boots)
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

            var bootcount = (double)boots.Count();
            foreach (KeyValuePair<string,int> pair in counts)
            {
                p_values.Add(pair.Key, (double)pair.Value / bootcount);
            }

            return p_values;
        }

        // Exclude specified input index, compute multinomial probabilty vector, and return true if probability is below threshold
        public static bool RejectNullHypothesis(FunctionOutput<string>[] boots, string original_output, int exclude_index, double significance)
        {
            // get bootstrap fingerprint for exclude_index
            var xfp = BigInteger.One << exclude_index;

            // filter bootstraps which include exclude_index
            var boots_exc = boots.Where(b => (b.GetExcludes() & xfp) == xfp);

            // get p_value vector
            var freq = BootstrapFrequency(boots_exc);

            // what is the probability of seeing the original output?
            double p_val;
            if (!freq.TryGetValue(original_output, out p_val))
            {
                p_val = 0.0;
            }

            // test H_0
            return p_val < 1.0 - significance;
        }

        // Exclude a specified input index, compute quantiles, and check position of original input
        public static double RejectNullHypothesis(FunctionOutput<double>[] boots, string original_output, int exclude_index, double significance)
        {
            // low
            double low_thresh = (1.0 - significance) / 2.0;

            // high
            double hi_thresh = significance + low_thresh;

            // get bootstrap fingerprint for exclude_index
            var xfp = BigInteger.One << exclude_index;

            // filter bootstraps that include exclude_index
            var boots_exc = boots.Where(b => (b.GetExcludes() & xfp) == xfp);

            var exc_count = boots_exc.Count();

            // return neutral (0.5) if we are having a sparsity problem
            if (exc_count == 0)
            {
                return 0.5;
            }
            // index for value greater than 2.5% of the lowest values; we want to round down here
            var low_index = System.Convert.ToInt32(Math.Floor((float)(exc_count - 1) * low_thresh));
            // index for value greater than 97.5% of the lowest values; we want to round up here
            var high_index = System.Convert.ToInt32(Math.Ceiling((float)(exc_count - 1) * hi_thresh));

            var low_value = boots_exc.ElementAt(low_index).GetValue();
            var high_value = boots_exc.ElementAt(high_index).GetValue();

            var lowest_value = boots_exc.ElementAt(0).GetValue();
            var highest_value = boots_exc.ElementAt(exc_count - 1).GetValue();

            double original_output_d;
            Double.TryParse(original_output, out original_output_d);

            // truncate the values to deal with floating point imprecision
            var low_value_tr = Math.Truncate(low_value * 10000) / 10000;
            var high_value_tr = Math.Truncate(high_value * 10000) / 10000;
            var original_tr = Math.Truncate(original_output_d * 10000) / 10000;
            
            var lowest_value_tr = Math.Truncate(lowest_value * 10000) / 10000;
            var highest_value_tr = Math.Truncate(highest_value * 10000) / 10000;

            // reject or fail to reject H_0
            if (original_tr > high_value_tr)
            {
                if (highest_value_tr != high_value_tr)
                {
                    return Math.Abs((original_tr - high_value_tr) / Math.Abs(high_value_tr - highest_value_tr)); //normalize by the highest 2.5%
                }
                else //can't normalize
                {
                    return Math.Abs(original_tr - high_value_tr);
                }
            }
            else if (original_tr < low_value_tr)
            {
                if (lowest_value_tr != low_value_tr)
                {
                    return Math.Abs((original_tr - low_value_tr) / Math.Abs(low_value_tr - lowest_value_tr));  //normalize by the lowest 2.5%
                }
                else //can't normalize
                {
                    return Math.Abs(original_tr - low_value_tr);
                }
            }

            return 0.0;
        }

        // Propagate weights
        private static void PropagateWeights(DAG dag)
        {
            if (dag.containsLoop())
            {
                throw new ContainsLoopException();
            }

            // starting set of functions; roots in the forest
            var formulas = dag.terminalFormulaNodes(false);

            // for each forest
            foreach (AST.Address f in formulas)
            {
                dag.setWeight(f, PropagateNodeWeight(f, dag));
            }
        }

        private static int PropagateNodeWeight(AST.Address node, DAG dag)
        {
            // if the node is a formula, recursively
            // compute its weight
            if (dag.isFormula(node))
            {
                // get input nodes
                var vector_rngs = dag.getFormulaInputVectors(node);
                var scinputs = dag.getFormulaSingleCellInputs(node);
                var inputs = vector_rngs.SelectMany(vrng => vrng.Addresses()).ToList();
                inputs.AddRange(scinputs);

                // call recursively and sum components
                var weight = 0;
                foreach (var input in inputs)
                {
                    weight += PropagateNodeWeight(input, dag);
                }
                dag.setWeight(node, weight);
                return weight;
            }
            // node is an input
            else
            {
                dag.setWeight(node, 1);
                return 1;
            }
        }
    }
}
