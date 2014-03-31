using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using TreeNode = DataDebugMethods.TreeNode;
using CellDict = System.Collections.Generic.Dictionary<AST.Address, string>;
using TreeScore = System.Collections.Generic.Dictionary<DataDebugMethods.TreeNode, int>;
using ErrorDict = System.Collections.Generic.Dictionary<AST.Address, double>;
using System.Diagnostics;
using DataDebugMethods;
using Serialization = System.Runtime.Serialization;
using OptChar = Microsoft.FSharp.Core.FSharpOption<char>;
using System.Runtime.InteropServices;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Formatters.Binary;
using System.IO;

namespace UserSimulation
{
    public class NoRangeInputs : Exception { }
    public class NoFormulas : Exception { }

    [Serializable]
    public class LogEntry
    {
        readonly AnalysisType _procedure;
        readonly string _filename;
        readonly AST.Address _address;
        readonly string _original_value;
        readonly string _erroneous_value;
        readonly double _output_error_magnitude;
        readonly double _input_error_magnitude;
        readonly bool _was_flagged;
        readonly bool _was_error;
        readonly double _significance;
        readonly double _threshold;
        public LogEntry(AnalysisType procedure,
                        string filename,
                        AST.Address address,
                        string original_value,
                        string erroneous_value,
                        double output_error_magnitude,
                        double input_error_magnitude,
                        bool was_flagged,
                        bool was_error,
                        double significance,
                        double threshold)
        {
            _procedure = procedure;
            _filename = filename;
            _address = address;
            _original_value = original_value;
            _erroneous_value = erroneous_value;
            _output_error_magnitude = output_error_magnitude;
            _input_error_magnitude = input_error_magnitude;
            _was_flagged = was_flagged;
            _was_error = was_error;
            _significance = significance;
            _threshold = threshold;
        }

        public static String Headers() {
            return "filename, procedure, significance, address, original_value, erroneous_value," +
                   "total_relative_error, typo_magnitude, was_flagged, was_error\n";
        }

        public void WriteLog(String logfile)
        {
            System.IO.File.AppendAllText(logfile, String.Format("{0},{1},{2},{3},{4},{5},{6},{7},{8},{9},{10}\n",
                                                        _filename,
                                                        _procedure,
                                                        _significance,
                                                        _threshold,
                                                        _address,
                                                        _original_value,
                                                        _erroneous_value,
                                                        _output_error_magnitude,
                                                        _input_error_magnitude,
                                                        _was_flagged,
                                                        _was_error));
        }
    }

    [Serializable]
    public enum ErrorCondition
    {
        OK,
        ContainsNoInputs,
        Exception
    }

    [Serializable]
    public enum AnalysisType
    {
        CheckCell           = 0,
        NormalPerRange      = 1,    //normal analysis of inputs on a per-range granularity
        NormalAllInputs     = 2     //normal analysis on the entire set of inputs
    }

    public class SimulationNotRunException : Exception
    {
        public SimulationNotRunException(string message) : base(message) { }
    } 

    [Serializable]
    public class Simulation
    {
        private bool _simulation_run = false;    // was the simulation run?
        private String _wb_name = "";
        private String _wb_path = "";
        private ErrorCondition _exit_state = ErrorCondition.OK;
        private string _exception_message = "";
        private UserResults _user;
        private ErrorDict _error;
        private double _total_relative_error = 0;
        private int _max_effort = 1;
        private int _cells_in_scope = 0;
        private int _effort = 0;
        private double _expended_effort = 0;
        private double _initial_total_relative_error = 0;
        private Dictionary<AST.Address, string> _errors = new Dictionary<AST.Address, string>();
        private double _average_precision = 0;
        private AnalysisType _analysis_type;
        private double _tree_construct_time = 0.0;
        private double _analysis_time = 0.0;
        private bool _normal_cutoff;
        private double _significance;
        private bool _all_outputs;
        private bool _weighted;
        private List<LogEntry> _error_log = new List<LogEntry>();

        public ErrorCondition GetExitState()
        {
            return _exit_state;
        }

        public List<AST.Address> GetTruePositives()
        {
            return _user.true_positives;
        }

        public List<AST.Address> GetFalsePositives()
        {
            return _user.false_positives;
        }

        public HashSet<AST.Address> GetFalseNegatives()
        {
            return _user.false_negatives;
        }

        public ErrorDict GetCalculatedError()
        {
            return _error;
        }

        public double GetTotalRelativeError()
        {
            return _total_relative_error;
        }

        public int GetMaxEffort()
        {
            return _max_effort;
        }

        public int GetEffort()
        {
            return _effort;
        }

        public double GetRelativeEffort()
        {
            return _expended_effort;
        }

        public void Serialize(string file_name)
        {
            IFormatter formatter = new BinaryFormatter();
            using (Stream stream = new FileStream(file_name, FileMode.Create, FileAccess.Write, FileShare.None))
            {
                formatter.Serialize(stream, this);
            }
        }

        public static Simulation Deserialize(string file_name)
        {
            Simulation sim;

            using (Stream stream = new FileStream(file_name, FileMode.Open, FileAccess.Read, FileShare.Read))
            {
                IFormatter formatter = new BinaryFormatter();
                sim = (Simulation)formatter.Deserialize(stream);
            }
            return sim;
        }

        // Get dictionary of inputs and the error they produce
        public Dictionary<AST.Address, Tuple<string, double>> TopOfKErrors(TreeNode[] terminal_formula_nodes, CellDict inputs, int k, CellDict correct_outputs, Excel.Application app, Excel.Workbook wb, string classification_file)
        {
            var eg = new ErrorGenerator();
            var c = Classification.Deserialize(classification_file);
            var max_error_produced_dictionary = new Dictionary<AST.Address, Tuple<string, double>>();

            foreach (KeyValuePair<AST.Address,string> pair in inputs)
            {
                AST.Address addr = pair.Key;
                string orig_value = pair.Value;

                //Load in the classification's dictionaries
                double max_error_produced = 0.0;
                string max_error_string = "";

                // get k strings, in parallel
                string[] errorstrings = eg.GenerateErrorStrings(orig_value, c, k);

                for (int i = 0; i < k; i++)
                {
                    CellDict cd = new CellDict();
                    cd.Add(addr, errorstrings[i]);
                    //inject the typo
                    InjectValues(app, wb, cd);

                    // save function outputs
                    CellDict incorrect_outputs = SaveOutputs(terminal_formula_nodes);

                    //remove the typo that was introduced
                    cd.Clear();
                    cd.Add(addr, orig_value);
                    InjectValues(app, wb, cd);

                    double total_error = CalculateTotalError(correct_outputs, incorrect_outputs);

                    //keep track of the largest observed max error
                    if (total_error > max_error_produced)
                    {
                        max_error_produced = total_error;
                        max_error_string = errorstrings[i];
                    }
                }
                //Add entry for this TreeNode in our dictionary with its max_error_produced
                max_error_produced_dictionary.Add(addr, new Tuple<string, double>(max_error_string, max_error_produced));
            }
            return max_error_produced_dictionary;
        }

        public CellDict GetTopErrors(Dictionary<AST.Address, Tuple<string, double>> max_error_produced_dictionary, double threshold)
        {
            int inputs_count = max_error_produced_dictionary.Count;
            CellDict top_errors = new CellDict();
            while ((top_errors.Count / (double)inputs_count) < threshold)
            {
                double max = 0.0;
                AST.Address max_addr = null;
                string max_node_string = "";
                //Find the max_node
                foreach (var kvp in max_error_produced_dictionary)
                {
                    if (kvp.Value.Item2 >= max)
                    {
                        max = kvp.Value.Item2;
                        max_addr = kvp.Key;
                        max_node_string = kvp.Value.Item1;
                    }
                }
                max_error_produced_dictionary.Remove(max_addr);
                top_errors.Add(max_addr, max_node_string);
            }

            return top_errors;
        }

        public void Run(int nboots,                 // number of bootstraps
                        string xlfile,              // name of the workbook
                        double significance,        // significance threshold for test
                        Excel.Application app,      // reference to Excel app
                        double threshold,           // percentage of erroneous cells
                        Classification c,           // data from which to generate errors
                        Random r,                   // a random number generator
                        AnalysisType analysisType,  // the type of analysis to run
                        bool weighted,              // should we weigh things?
                        bool all_outputs,           // if !all_outputs, we only consider terminal outputs
                        bool normal_cutoff,         // indicates if we should use normal cutoff or top 5% for errors
                        AnalysisData data,
                        Excel.Workbook wb,
                        TreeNode[] terminal_formula_nodes,
                        TreeNode[] terminal_input_nodes,
                        CellDict original_inputs,
                        CellDict correct_outputs,
                        long max_duration_in_ms,
                        String logfile              //filename for the output log
                       )
        {
            //set wbname and path
            _wb_name = xlfile;
            _wb_path = wb.Path; 
            _analysis_type = analysisType;
            _normal_cutoff = normal_cutoff;
            _significance = significance;
            _all_outputs = all_outputs;
            _weighted = weighted;

            //Now we want to inject the errors from top_errors
            InjectValues(app, wb, _errors);

            // TODO: save a copy of the workbook for later inspection

            // save function outputs
            CellDict incorrect_outputs = SaveOutputs(terminal_formula_nodes);

            //Time the removal of errors
            Stopwatch sw = new Stopwatch();
            sw.Start();

            // remove errors until none remain; MODIFIES WORKBOOK
            _user = SimulateUser(nboots, significance, threshold, data, original_inputs, _errors, correct_outputs, wb, app, analysisType, weighted, all_outputs, normal_cutoff, max_duration_in_ms, sw, logfile);

            sw.Stop();
            TimeSpan elapsed = sw.Elapsed;
            _analysis_time = elapsed.TotalSeconds;

            // save partially-corrected outputs
            var partially_corrected_outputs = SaveOutputs(terminal_formula_nodes);

            // compute total relative error
            _error = CalculateNormalizedError(correct_outputs, partially_corrected_outputs, _user.max_errors);
            _total_relative_error = TotalRelativeError(_error);

            // compute starting total relative error (normalized by max_errors)
            ErrorDict starting_error = CalculateNormalizedError(correct_outputs, incorrect_outputs, _user.max_errors);
            _initial_total_relative_error = TotalRelativeError(starting_error);

            // effort
            _max_effort = data.cell_nodes.Count;
            _cells_in_scope = 0;

            foreach (TreeNode input_range in terminal_input_nodes)
            {
                foreach (TreeNode input_to_range in input_range.getInputs())
                {
                    if (input_to_range.isCell() && !input_to_range.isFormula()) //if this input is a cell and is not a formula, then it is perturbable, so it's in our scope
                    {
                        _cells_in_scope++;
                    }
                }
            }

            _effort = (_user.true_positives.Count + _user.false_positives.Count);
            _expended_effort = (double)_effort / (double)_max_effort;

            // compute average precision
            // AveP = (\sum_{k=1}^n (P(k) * rel(k))) / |total positives|
            // where P(k) is the precision at threshold k,
            // rel(k) = \{ 1 if item at k is a true positive, 0 otherwise
            _average_precision = _user.PrecRel_at_k.Sum() / (double)_errors.Count;

            // restore original values
            InjectValues(app, wb, original_inputs);

            _tree_construct_time = data.tree_construct_time;
            // flag that we're done; safe to print output results
            _simulation_run = true;
        }

        // create and run a CheckCell simulation
        public void Run(int nboots,                 // number of bootstraps
                        string xlfile,              // name of the workbook
                        double significance,        // significance threshold for test
                        Excel.Application app,      // reference to Excel app
                        double threshold,           // percentage of erroneous cells
                        Classification c,           // data from which to generate errors
                        Random r,                   // a random number generator
                        AnalysisType analysisType,  // the type of analysis to run
                        bool weighted,              // should we weigh things?
                        bool all_outputs,           // if !all_outputs, we only consider terminal outputs
                        bool normal_cutoff,         // indicates if we should use a normal cutoff or top x%
                        long max_duration_in_ms,    // maximum duration before throwing a timeout exception
                        String logfile              //filename for the output log
                       )
        {
            // open workbook
            Excel.Workbook wb = Utility.OpenWorkbook(xlfile, app);

            // build dependency graph
            var data = ConstructTree.constructTree(app.ActiveWorkbook, app);

            // create ErrorGenerator object
            var egen = new ErrorGenerator();

            // get terminal input and terminal formula nodes once
            var terminal_input_nodes = data.TerminalInputNodes();
            var terminal_formula_nodes = data.TerminalFormulaNodes(all_outputs);

            if (terminal_input_nodes.Length == 0)
            {
                throw new NoRangeInputs();
            }

            // save original spreadsheet state
            CellDict original_inputs = SaveInputs(terminal_input_nodes);
            if (original_inputs.Count() == 0)
            {
                throw new NoFormulas();
            }

            // force a recalculation before saving outputs, otherwise we may
            // erroneously conclude that the procedure did the wrong thing
            // based solely on Excel floating-point oddities
            InjectValues(app, wb, original_inputs);

            // save function outputs
            CellDict correct_outputs = SaveOutputs(terminal_formula_nodes);

            // generate errors
            _errors = egen.RandomlyGenerateErrors(original_inputs, c, threshold);

            Run(nboots, xlfile, significance, app, threshold, c, r, analysisType, weighted, all_outputs, normal_cutoff, data, wb, terminal_formula_nodes, terminal_input_nodes, original_inputs, correct_outputs, max_duration_in_ms, logfile);
        }

        // For running a simulation from the batch runner
        public void RunFromBatch(int nboots,        // number of bootstraps
                        string xlfile,              // name of the workbook
                        double significance,        // significance threshold for test
                        Excel.Application app,      // reference to Excel app
                        double threshold,           // percentage of erroneous cells
                        Classification c,           // data from which to generate errors
                        Random r,                   // a random number generator
                        AnalysisType analysisType,  // the type of analysis to run
                        bool weighted,              // should we weigh things?
                        bool all_outputs,           // if !all_outputs, we only consider terminal outputs
                        bool normal_cutoff,         // indicates if we should use a normal cutoff or top x%
                        AnalysisData data,          // the computation tree of the spreadsheet
                        Excel.Workbook wb,          // the workbook being analyzed
                        CellDict errors,            // the errors that will be introduced in the spreadsheet
                        TreeNode[] terminal_input_nodes,   // the inputs
                        TreeNode[] terminal_formula_nodes, // the outputs
                        CellDict original_inputs,          // original values of the inputs
                        CellDict correct_outputs,          // the correct outputs
                        long max_duration_in_ms,
                        String logfile              //filename for the output log
                       )
        {
            if (terminal_input_nodes.Length == 0)
            {
                throw new NoRangeInputs();
            }

            if (original_inputs.Count() == 0)
            {
                throw new NoFormulas();
            }

            _errors = errors;
                
            Run(nboots, xlfile, significance, app, threshold, c, r, analysisType, weighted, all_outputs, normal_cutoff, data, wb, terminal_formula_nodes, terminal_input_nodes, original_inputs, correct_outputs, max_duration_in_ms, logfile);
        }

        public double RemainingError()
        {
            return _total_relative_error / _initial_total_relative_error;
        }

        public static String HeaderRowForCSV()
        {
            return String.Format("{0},{1},{2},{3},{4},{5},{6},{7},{8},{9},{10},{11},{12},{13},{14},{15},{16},{17},{18},{19},{20}",
                                 "workbook_name",                               //0
                                 "initial_total_relative_error",                //1
                                 "total_relative_error",                        //2
                                 "remaining_error",                             //3
                                 "effort",                                      //4
                                 "max_effort",                                  //5
                                 "cells_in_scope",                              //6
                                 "ratio_scope_out_of_total",                    //7
                                 "expended_effort",                             //8
                                 "number_of_errors",                            //9
                                 "true_positives",                              //10
                                 "false_positives",                             //11
                                 "false_negatives",                             //12
                                 "average_precision",                           //13
                                 "tree_construct_time_seconds",                 //14
                                 "bootstrap_or_normal_time_seconds",            //15
                                 "analysis_type",                               //16
                                 "normal_cutoff",                               //17
                                 "significance",                                //18
                                 "all_outputs",                                 //19
                                 "weighted\n");                                 //20
        }

        public String FormatResultsAsCSV()
        {
            return _wb_name + "," +                         // workbook name
                    _initial_total_relative_error + "," +   // initial total relative error
                    _total_relative_error + "," +           // final total relative error
                    RemainingError() + "," +                // remaining error
                    _effort.ToString() + "," +              // effort
                    _max_effort + "," +                     // max effort
                    _cells_in_scope + "," +                 // perturbable cells (these are in our scope)
                    (double)_cells_in_scope / (double)_max_effort + "," +     // proportion of cells that are in scopes of our tool
                    _expended_effort + "," +                // expended effort
                    _errors.Count + "," +                   // number of errors
                    _user.true_positives.Count + "," +      // number of true positives
                    _user.false_positives.Count + "," +     // number of false positives
                    _user.false_negatives.Count + "," +     // number of false negatives
                    _average_precision + "," +              // average precision
                    _tree_construct_time + "," +            // tree construction time in seconds
                    _analysis_time + "," +                  // bootstrap or normal analysis time in seconds
                    _analysis_type + "," +                  // anaysis type (CheckCell, normal per range, normal on all inputs)
                    _normal_cutoff + "," +
                    _significance + "," +
                    _all_outputs + "," +
                    _weighted + "\n";
        }

        //This method creates a csv file that shows the error reduction after each fix is applied
        public static void ToTimeseriesCSV(Excel.Workbook wb, double current_error, double current_effort)
        {
            //The file (timeseries_results.csv) is created in the same directory as the file currently being analyzed
            string dir_path = wb.Path;
            string file_path = dir_path + "\\timeseries_results.csv";
            string text = "";

            //if file exists, read existing data
            if (System.IO.File.Exists(file_path))
            {
                text = System.IO.File.ReadAllText(file_path);
                text += "\n";
            }
            //otherwise write header
            else
            {
                text = "workbook_name,current_error,current_effort\n";
            }

            text += String.Format("{0},{1},{2}\n", wb.Name, current_error, current_effort);

            System.IO.File.WriteAllText(file_path, text);
        }

        /// <summary>
        /// Creates a CSV file with information about the CheckCell oracle simulation.
        /// </summary>
        /// <param name="output_filename">Path for writing CSV.</param>
        public void ToCSV(string output_filename)
        {
            if (_simulation_run)
            {
                // the file is created at the following path
                // if file exists, read it and append to it
                if (System.IO.File.Exists(output_filename))
                {
                    string text = System.IO.File.ReadAllText(output_filename);
                    text += "\n" + FormatResultsAsCSV();
                    System.IO.File.WriteAllText(output_filename, text);
                }
                // otherwise create the file, adding the column headers, and write to it
                else
                {
                    System.IO.File.WriteAllText(output_filename, HeaderRowForCSV() + "\n" + FormatResultsAsCSV());
                }
            }
            else
            {
                throw new SimulationNotRunException(_exception_message);
            }
        }

        private static void UpdatePerFunctionMaxError(CellDict correct_outputs, CellDict incorrect_outputs, ErrorDict max_errors)
        {
            // for each output
            foreach (var kvp in correct_outputs)
            {
                var addr = kvp.Key;
                var correct_value = correct_outputs[addr];
                var incorrect_value = incorrect_outputs[addr];
                // numeric
                if (ExcelParser.isNumeric(correct_value) && ExcelParser.isNumeric(incorrect_value))
                {
                    var error = Math.Abs(System.Convert.ToDouble(correct_value) - System.Convert.ToDouble(incorrect_value));
                    if (max_errors.ContainsKey(addr))
                    {
                        if (max_errors[addr] < error)
                        {
                            max_errors[addr] = error;
                        }
                    }
                    else
                    {
                        max_errors.Add(addr, error);
                    }
                }
                // non-numeric
                else
                {
                    var error = correct_value.Equals(incorrect_value) ? 0.0 : 1.0;
                    if (max_errors.ContainsKey(addr))
                    {
                        if (error > 0)
                        {
                            max_errors[addr] = error;
                        }
                    }
                    else
                    {
                        max_errors.Add(addr, error);
                    }
                }
            }
        }

        //The total error is the sum of the absolute errors of all outputs
        private static double CalculateTotalError(CellDict correct_outputs, CellDict incorrect_outputs)
        {
            //Iterate over all outputs and accumulate the total error
            double total_error = 0.0;
            foreach (var kvp in correct_outputs)
            {
                var addr = kvp.Key;
                var correct_value = correct_outputs[addr];
                var incorrect_value = incorrect_outputs[addr];
                if (ExcelParser.isNumeric(correct_value) && ExcelParser.isNumeric(incorrect_value))
                {
                    total_error += Math.Abs(System.Convert.ToDouble(correct_value) - System.Convert.ToDouble(incorrect_value));
                }
                else
                {
                    total_error += correct_value.Equals(incorrect_value) ? 0.0 : 1.0;
                }
            }
            return total_error;
        }

        [Serializable]
        private class UserResults
        {
            public List<AST.Address> true_positives = new List<AST.Address>();
            public List<AST.Address> false_positives = new List<AST.Address>();
            public HashSet<AST.Address> false_negatives = new HashSet<AST.Address>();
            //Keeps track of the largest errors we observe during the simulation for each output
            public ErrorDict max_errors = new ErrorDict();
            public List<double> current_total_error = new List<double>();
            public List<double> PrecRel_at_k = new List<double>();
        }

        //Computes total relative error
        //Each entry in the dictionary is normalized to its max value, so they are all <= 1.0.
        //We sum them up and divide by the total number of entries to get the total relative error
        private static double TotalRelativeError(ErrorDict error)
        {
            return error.Select(pair => pair.Value).Sum() / (double)error.Count();
        }

        private static ErrorDict CalculateNormalizedError(CellDict correct_outputs, CellDict partially_corrected_outputs, ErrorDict max_errors)
        {
            var ed = new ErrorDict();

            foreach (KeyValuePair<AST.Address, string> orig in correct_outputs)
            {
                var addr = orig.Key;
                string correct_value = orig.Value;
                string partially_corrected_value = System.Convert.ToString(partially_corrected_outputs[addr]);
                // if the function produces numeric outputs, calculate distance
                if (ExcelParser.isNumeric(correct_value) &&
                    ExcelParser.isNumeric(partially_corrected_value))
                {
                    ed.Add(addr, RelativeNumericError(System.Convert.ToDouble(correct_value),
                                                      System.Convert.ToDouble(partially_corrected_value),
                                                      max_errors[addr]));
                }
                // calculate indicator function
                else
                {
                    ed.Add(addr, RelativeCategoricalError(correct_value, partially_corrected_value));
                }
            }

            return ed;
        }

        // compares the corrected function output against the incorrected output
        // 0 means that the error has been completely corrected; 1 means that
        // the error totally remains
        private static double RelativeNumericError(double correct_value, double partially_corrected_value, double max_error)
        {
            //|f(I'') - f(I)| / max f(I')
            if (max_error != 0.0)
            {
                return Math.Abs(partially_corrected_value - correct_value) / max_error;
            }
            else
            {
                return 0;
            }
        }

        private static double RelativeCategoricalError(string original_value, string partially_corrected_value)
        {
            if (String.Equals(original_value, partially_corrected_value))
            {
                return 0;
            } else {
                return 1;
            }
        }

        private static AST.Address CheckCell_Step(UserResults o,
                                           double significance,
                                           double threshold,
                                           int nboots,
                                           AnalysisData data,
                                           Excel.Application app,
                                           bool weighted,
                                           bool all_outputs,
                                           bool run_bootstrap,
                                           bool normal_cutoff,
                                           HashSet<AST.Address> known_good,
                                           ref List<KeyValuePair<TreeNode, int>> filtered_high_scores,
                                           long max_duration_in_ms,
                                           Stopwatch sw)
        {
            // Get bootstraps
            // The bootstrap should only re-run if there is a correction made, 
            //      not when something is marked as OK (isn't one of the introduced errors)
            // The list of suspected cells doesn't change when we mark something as OK,
            //      we just move on to the next thing in the list
            if (run_bootstrap)
            {
                TreeScore scores = Analysis.Bootstrap(nboots, data, app, weighted, all_outputs, max_duration_in_ms, sw, significance);
                var scores_list = scores.OrderByDescending(pair => pair.Value).ToList();

                //Using an outlier test for highlighting 
                //scores that fall outside of two standard deviations from the others
                //The one-sided 5% cutoff for the normal distribution is 1.6448.

                if (normal_cutoff)
                {
                    //Code for doing normal outlier analysis on the scores:
                    //find mean:
                    double sum = 0.0;
                    foreach (double d in scores.Values)
                    {
                        sum += d;
                    }
                    double mean = sum / scores.Values.Count;
                    //find variance
                    double distance_sum_sq = 0.0;
                    foreach (double d in scores.Values)
                    {
                        distance_sum_sq += Math.Pow(mean - d, 2);
                    }
                    double variance = distance_sum_sq / scores.Values.Count;

                    //find std. deviation
                    double std_deviation = Math.Sqrt(variance);

                    if (threshold == 0.05)
                    {
                        filtered_high_scores = scores_list.Where(kvp => kvp.Value > mean + std_deviation * 1.6448).ToList();
                    }
                    else if (threshold == 0.1)   //10% cutoff 1.2815
                    {
                        filtered_high_scores = scores_list.Where(kvp => kvp.Value > mean + std_deviation * 1.2815).ToList();
                    }
                    else if (threshold == 0.025) //2.5% cutoff 1.9599
                    {
                        filtered_high_scores = scores_list.Where(kvp => kvp.Value > mean + std_deviation * 1.9599).ToList();
                    }
                    else if (threshold == 0.075) //7.5% cutoff 1.4395
                    {
                        filtered_high_scores = scores_list.Where(kvp => kvp.Value > mean + std_deviation * 1.4395).ToList();
                    }
                    else
                    {
                        throw new Exception("Uhhh.... What's my cutoff?");
                    }
                }
                else
                {

                    int start_ptr = 0;
                    int end_ptr = 0;

                    List<KeyValuePair<TreeNode, int>> high_scores = new List<KeyValuePair<TreeNode, int>>();

                    while ((double)start_ptr / scores_list.Count < threshold) //the start of this score region is before the cutoff
                    {
                        //while the scores at the start and end pointers are the same, bump the end pointer
                        while (end_ptr < scores_list.Count && scores_list[start_ptr].Value == scores_list[end_ptr].Value)
                        {
                            end_ptr++;
                        }
                        //Now the end_pointer points to the first index with a lower score
                        //If the number of entries with the current value is fewer than the significance cutoff, add all values of this score to the high_scores list; the number of entries is equal to the end_ptr since end_ptr is zero-based
                        //There is some added "wiggle room" to the cutoff, so that the last entry is allowed to straddle the cutoff bound.
                        //  To do this, we add (1 / total number of entries) to the cutoff
                        //The purpose of the wiggle room is to allow us to deal with small ranges (less than 20 entries), since a single entry accounts
                        //for more than 5% of the total.
                        if ((double)end_ptr / scores_list.Count < threshold + (double)1.0 / scores_list.Count)
                        {
                            //add all values of the current score to high_scores list
                            for (; start_ptr < end_ptr; start_ptr++)
                            {
                                high_scores.Add(scores_list[start_ptr]);
                            }
                            //Increment the start pointer to the start of the next score region
                            start_ptr++;
                        }
                        else    //if this score region extends past the cutoff, we don't add any of its values to the high_scores list, and stop
                        {
                            break;
                        }
                    }
                    // filter out cells marked as OK
                    filtered_high_scores = high_scores.Where(kvp => !known_good.Contains(kvp.Key.GetAddress())).ToList();
                }
            }
            else  //if no corrections were made (a cell was marked as OK, not corrected)
            {
                //re-filter out cells marked as OK
                filtered_high_scores = filtered_high_scores.Where(kvp => !known_good.Contains(kvp.Key.GetAddress())).ToList();
            }
            if (filtered_high_scores.Count() != 0)
            {
                // get TreeNode corresponding to most unusual score
                return filtered_high_scores[0].Key.GetAddress();
            }
            else
            {
                return null;
            }
        }

        private static AST.Address NormalPerRange_Step(AnalysisData data,
                                                      Excel.Workbook wb,
                                                      HashSet<AST.Address> known_good,
                                                      long max_duration_in_ms,
                                                      Stopwatch sw)
        {
            AST.Address flagged_cell = null;

            //Generate normal distributions for every input range until an error is found
            //Then break out of the loop and report it.
            foreach (var range in data.input_ranges.Values)
            {
                var normal_dist = new DataDebugMethods.NormalDistribution(range.getCOMObject());

                // Get top outlier which has not been inspected already
                if (normal_dist.errorsCount() > 0)
                {
                    for (int i = 0; i < normal_dist.errorsCount(); i++)
                    {
                        // check for timeout
                        if (sw.ElapsedMilliseconds > max_duration_in_ms)
                        {
                            throw new TimeoutException("Timeout exception in NormalPerRange_Step.");
                        }

                        var flagged_com = normal_dist.getError(i);
                        flagged_cell = (new TreeNode(flagged_com, flagged_com.Worksheet, wb)).GetAddress();
                        if (known_good.Contains(flagged_cell))
                        {
                            flagged_cell = null;
                        }
                        else
                        {
                            break;
                        }
                    }
                }
                //If a cell is flagged, do not move on to the next range (if you do, you'll overwrite the flagged_cell
                if (flagged_cell != null)
                {
                    break;
                }
            }

            return flagged_cell;
        }

        private static AST.Address NormalAllOutputs_Step(AnalysisData data,
                                                         Excel.Application app,
                                                         Excel.Workbook wb,
                                                         HashSet<AST.Address> known_good,
                                                         long max_duration_in_ms,
                                                         Stopwatch sw)
        {
            AST.Address flagged_cell = null;

            //Generate a normal distribution for the entire set of inputs
            var normal_dist = new DataDebugMethods.NormalDistribution(data.TerminalInputNodes(), app);

            // Get top outlier
            if (normal_dist.errorsCount() > 0)
            {
                for (int i = 0; i < normal_dist.errorsCount(); i++)
                {
                    // check for timeout
                    if (sw.ElapsedMilliseconds > max_duration_in_ms)
                    {
                        throw new TimeoutException("Timeout exception in NormalAllOutputs_Step.");
                    }

                    var flagged_com = normal_dist.getError(i);
                    flagged_cell = (new TreeNode(flagged_com, flagged_com.Worksheet, wb)).GetAddress();
                    if (known_good.Contains(flagged_cell))
                    {
                        flagged_cell = null;
                    }
                    else
                    {
                        break;
                    }
                }
            }

            return flagged_cell;
        }

        private double InputErrorMagnitude(string error, string correct)
        {
            double e, c;
            if (Double.TryParse(error, out e) && Double.TryParse(correct, out c))
            {
                if (c != 0)
                {
                    return Math.Abs(e / c);
                }
            }
            return 0;
        }

        private double MeanErrorMagnitude(CellDict partially_corrected_outputs, CellDict original_outputs)
        {
            int count = 0;
            double magnitude = 0;

            foreach (KeyValuePair<AST.Address, string> pair in partially_corrected_outputs)
            {
                var err = pair.Value;
                var cor = original_outputs[pair.Key];

                // if the denominator is a string, do nothing
                double c;
                double e;
                if (Double.TryParse(cor, out c))
                {
                    //if the error is an empty string, convert it to a 0
                    if (String.IsNullOrWhiteSpace(err))
                    {
                        e = 0.0;
                    }
                    //if the error is a number, get its value
                    else if (Double.TryParse(err, out e)) { }
                    //for all other strings, do nothing
                    else 
                    {
                        continue;
                    }
                    count++;
                    magnitude += Math.Abs(e - c) / Math.Abs(c);
                }
            }

            if (count == 0)
            {
                return 0.0;
            }
            return magnitude / (double)count;
        }

        // remove errors until none remain
        private UserResults SimulateUser(int nboots,
                                         double significance,
                                         double threshold,
                                         AnalysisData data,
                                         CellDict original_inputs,
                                         CellDict errord,
                                         CellDict correct_outputs,
                                         Excel.Workbook wb,
                                         Excel.Application app,
                                         AnalysisType analysis_type,
                                         bool weighted,
                                         bool all_outputs,
                                         bool normal_cutoff,
                                         long max_duration_in_ms,
                                         Stopwatch sw,
                                         String logfile
                                        )
        {
            // init user results data structure
            var o = new UserResults();
            HashSet<AST.Address> known_good = new HashSet<AST.Address>();

            // initialize procedure
            var errors_remain = true;
            var max_errors = new ErrorDict();
            var incorrect_outputs = SaveOutputs(data.TerminalFormulaNodes(all_outputs));
            var errors_found = 0;
            var number_of_true_errors = errord.Count;
            UpdatePerFunctionMaxError(correct_outputs, incorrect_outputs, max_errors);

            // the corrected state of the spreadsheet
            CellDict partially_corrected_outputs = correct_outputs.ToDictionary(p => p.Key, p => p.Value);

            // remove errors loop
            var cells_inspected = 0;
            List<KeyValuePair<TreeNode, int>> filtered_high_scores = null;
            bool correction_made = true;
            while (errors_remain)
            {
                Console.Write(".");

                AST.Address flagged_cell = null;

                // choose the appropriate test
                // TODO: the test type really should be a lambda
                if (analysis_type == AnalysisType.CheckCell)
                {
                    flagged_cell = CheckCell_Step(o,
                                                  significance,
                                                  threshold,
                                                  nboots,
                                                  data,
                                                  app,
                                                  weighted,
                                                  all_outputs,
                                                  correction_made,
                                                  normal_cutoff,
                                                  known_good,
                                                  ref filtered_high_scores,
                                                  max_duration_in_ms,
                                                  sw);
                } else if (analysis_type == AnalysisType.NormalPerRange)
                {
                    flagged_cell = NormalPerRange_Step(data, wb, known_good, max_duration_in_ms, sw);
                }
                else if (analysis_type == AnalysisType.NormalAllInputs)
                {
                    flagged_cell = NormalAllOutputs_Step(data, app, wb, known_good, max_duration_in_ms, sw);
                }

                if (flagged_cell == null)
                {
                    errors_remain = false;
                }
                else    // a cell was flagged
                {
                    //cells_inspected should only be incremented when a cell is actually flagged. If nothing is flagged, 
                    //then nothing is inspected, so cells_inspected doesn't increase.
                    cells_inspected += 1;

                    // check to see if the flagged value is actually an error
                    if (errord.ContainsKey(flagged_cell))
                    {
                        correction_made = true;
                        errors_found += 1;
                        // P(k) * rel(k)
                        o.PrecRel_at_k.Add(errors_found / (double)cells_inspected);
                        o.true_positives.Add(flagged_cell);

                        // correct flagged cell
                        flagged_cell.GetCOMObject(app).Value2 = original_inputs[flagged_cell];
                        
                        UpdatePerFunctionMaxError(correct_outputs, partially_corrected_outputs, max_errors);
                        
                        // compute total error after applying this correction
                        var current_total_error = CalculateTotalError(correct_outputs, partially_corrected_outputs);
                        o.current_total_error.Add(current_total_error);

                        // save outputs
                        partially_corrected_outputs = SaveOutputs(data.TerminalFormulaNodes(all_outputs));
                    }
                    else
                    {
                        correction_made = false;
                        // numerator is 0 here because rel(k) = 0 when no error was found
                        o.PrecRel_at_k.Add(0.0);
                        o.false_positives.Add(flagged_cell);
                    }

                    // mark it as known good -- at this point the cell has been 
                    //      'inspected' regardless of whether it was an error
                    //      It was either corrected or marked as OK
                    known_good.Add(flagged_cell);

                    // compute output error magnitudes
                    var output_error_magnitude = MeanErrorMagnitude(partially_corrected_outputs, correct_outputs);
                    // compute input error magnitude
                    double input_error_magnitude;
                    if (errord.ContainsKey(flagged_cell))
                    {
                        input_error_magnitude = InputErrorMagnitude(errord[flagged_cell], original_inputs[flagged_cell]);
                    }
                    else
                    {
                        input_error_magnitude = 1.0;
                    }

                    // write error log
                    var logentry = new LogEntry(analysis_type,
                                                wb.Name,
                                                flagged_cell,
                                                original_inputs[flagged_cell],
                                                errord.ContainsKey(flagged_cell) ? errord[flagged_cell] : original_inputs[flagged_cell],
                                                output_error_magnitude,
                                                input_error_magnitude,
                                                true,
                                                correction_made,
                                                significance,
                                                threshold);
                    logentry.WriteLog(logfile);
                    _error_log.Add(logentry);
                }
            }

            // find all of the false negatives
            o.false_negatives = GetFalseNegatives(o.true_positives, o.false_positives, errord);
            o.max_errors = max_errors;

            var last_out_err_mag = MeanErrorMagnitude(partially_corrected_outputs, correct_outputs);

            // write out all false negative information
            foreach (AST.Address fn in o.false_negatives)
            {
                // write error log
                _error_log.Add(new LogEntry(analysis_type,
                                            wb.Name,
                                            fn,
                                            original_inputs[fn],
                                            errord[fn],
                                            last_out_err_mag,
                                            InputErrorMagnitude(errord[fn], original_inputs[fn]),
                                            false,
                                            true,
                                            significance,
                                            threshold));
            }
            return o;
        }

        private static CellDict ErrorDBToCellDict(ErrorDB errors)
        {
            var d = new CellDict();
            foreach (Error e in errors.Errors)
            {
                d.Add(e.GetAddress(), e.value);
            }
            return d;
        }

        // return the set of false negatives
        private static HashSet<AST.Address> GetFalseNegatives(List<AST.Address> true_positives, List<AST.Address> false_positives, CellDict errors)
        {
            var fnset = new HashSet<AST.Address>();
            var tpset = new HashSet<AST.Address>(true_positives);
            var fpset = new HashSet<AST.Address>(false_positives);

            foreach(KeyValuePair<AST.Address, string> pair in errors)
            {
                var addr = pair.Key;
                if (!tpset.Contains(addr) && !fpset.Contains(addr))
                {
                    fnset.Add(addr);
                }
            }

            return fnset;
        }

        // save spreadsheet inputs to a CellDict
        public static CellDict SaveInputs(TreeNode[] input_ranges)
        {
            try
            {
                var cd = new CellDict();
                foreach (TreeNode input_range in input_ranges)
                {
                    foreach (TreeNode cell in input_range.getInputs())
                    {
                        // never save formula; there's no point since we don't perturb them
                        var comcell = cell.getCOMObject();
                        var addr = cell.GetAddress();
                        if (!cd.ContainsKey(addr))
                        {
                            cd.Add(addr, cell.getCOMValueAsString());
                        }
                    }
                }
                return cd;
            }
            catch (Exception e)
            {
                throw new Exception(String.Format("Failed in SaveInputs: {0}", e.Message));
            }
        }

        // save spreadsheet outputs to a CellDict
        public static CellDict SaveOutputs(TreeNode[] formula_nodes)
        {
            var cd = new CellDict();
            foreach (TreeNode formula_cell in formula_nodes)
            {
                // throw an exception in debug mode, because this should never happen
                if (!(bool)formula_cell.getCOMObject().HasFormula)
                {
                    throw new Exception("Formula TreeNode has no formula.");
                }

                // get address
                var addr = formula_cell.GetAddress();

                // save value
                if (cd.ContainsKey(addr))
                {
                    throw new Exception(String.Format("Failed in SaveOutputs."));
                } else {
                    cd.Add(addr, formula_cell.getCOMValueAsString());
                }
            }
            return cd;
        }

        // inject errors into a workbook
        public static void InjectValues(Excel.Application app, Excel.Workbook wb, CellDict values)
        {
            foreach (KeyValuePair<AST.Address,string> pair in values)
            {
                var addr = pair.Key;
                var errorstr = pair.Value;
                var comcell = addr.GetCOMObject(app);

                // never perturb formulae
                if (!comcell.HasFormula)
                {
                    // inject error
                    comcell.Value2 = errorstr;
                }
            }
        }

        // Get dictionary of inputs and the error they produce
        public static CellDict GenImportantErrors(TreeNode[] output_nodes,
                                                  CellDict inputs,
                                                  int k,         // number of alternatives to consider
                                                  CellDict correct_outputs,
                                                  Excel.Application app, 
                                                  Excel.Workbook wb, 
                                                  Classification c)
        {
            var eg = new ErrorGenerator();
            var max_error_produced_dictionary = new Dictionary<AST.Address, Tuple<string, double>>();

            foreach (KeyValuePair<AST.Address, string> pair in inputs)
            {
                AST.Address addr = pair.Key;
                string orig_value = pair.Value;

                //Load in the classification's dictionaries
                double max_error_produced = 0.0;
                string max_error_string = "";

                // get k strings
                string[] errorstrings = eg.GenerateErrorStrings(orig_value, c, k);

                for (int i = 0; i < k; i++)
                {
                    CellDict cd = new CellDict();
                    cd.Add(addr, errorstrings[i]);
                    //inject the typo
                    InjectValues(app, wb, cd);

                    // save function outputs
                    CellDict incorrect_outputs = SaveOutputs(output_nodes);

                    //remove the typo that was introduced
                    cd.Clear();
                    cd.Add(addr, orig_value);
                    InjectValues(app, wb, cd);

                    double total_error = CalculateTotalError(correct_outputs, incorrect_outputs);

                    //keep track of the largest observed max error
                    if (total_error > max_error_produced)
                    {
                        max_error_produced = total_error;
                        max_error_string = errorstrings[i];
                    }
                }
                //Add entry for this TreeNode in our dictionary with its max_error_produced
                max_error_produced_dictionary.Add(addr, new Tuple<string, double>(max_error_string, max_error_produced));
            }

            // sort by max_error_produced
            var maxen = max_error_produced_dictionary.OrderByDescending(pair => pair.Value.Item2).Select(pair => new Tuple<AST.Address, string>(pair.Key, pair.Value.Item1)).ToList();

            return maxen.Take((int)Math.Ceiling(0.05 * inputs.Count)).ToDictionary(tup => tup.Item1, tup => tup.Item2);
        }

        public byte[] Serialize()
        {
            byte[] data;
            using (var ms = new System.IO.MemoryStream())
            {
                var formatter = new Serialization.Formatters.Binary.BinaryFormatter();
                formatter.Serialize(ms, this);
                data = ms.ToArray();
            }
            return data;
        }

        public static Simulation Deserialize(byte[] data)
        {
            Simulation s;
            using (var ms = new System.IO.MemoryStream())
            {
                ms.Read(data, 0, data.Length);
                var formatter = new Serialization.Formatters.Binary.BinaryFormatter();
                s = (Simulation)formatter.Deserialize(ms);
            }
            return s;
        }
    }
}
