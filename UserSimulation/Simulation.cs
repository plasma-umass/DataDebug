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
using OptString = Microsoft.FSharp.Core.FSharpOption<string>;
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
        readonly double _num_input_error_magnitude;
        readonly double _str_input_error_magnitude;
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
                        double num_input_error_magnitude,
                        double str_input_error_magnitude,
                        bool was_flagged,
                        bool was_error,
                        double significance,
                        double threshold)
        {
            _filename = filename;
            _procedure = procedure;
            _address = address;
            _original_value = original_value;
            _erroneous_value = erroneous_value;
            _output_error_magnitude = output_error_magnitude;
            _num_input_error_magnitude = num_input_error_magnitude;
            _str_input_error_magnitude = str_input_error_magnitude;
            _was_flagged = was_flagged;
            _was_error = was_error;
            _significance = significance;
            _threshold = threshold;
        }

        public static String Headers() {
            return "filename, " + // 0
                   "procedure, " + // 1
                   "significance, " + // 2
                   "threshold, " + // 3
                   "address, " + // 4
                   "original_value, " + // 5
                   "erroneous_value," + // 6
                   "total_relative_error, " + // 7
                   "num_input_err_mag, " + // 8
                   "str_input_err_mag, " + // 9
                   "was_flagged, " + // 10
                   "was_error" + // 11
                   Environment.NewLine; // 12
        }

        public void WriteLog(String logfile)
        {
            if (!System.IO.File.Exists(logfile))
            {
                System.IO.File.AppendAllText(logfile, Headers());
            }
            System.IO.File.AppendAllText(logfile, String.Format("{0},{1},{2},{3},{4},{5},{6},{7},{8},{9},{10},{11}{12}",
                                                        _filename, // 0
                                                        _procedure, // 1
                                                        _significance, // 2
                                                        _threshold, // 3
                                                        _address.A1Local(), // 4
                                                        _original_value, // 5
                                                        _erroneous_value, // 6
                                                        _output_error_magnitude,// 7
                                                        _num_input_error_magnitude, // 8
                                                        _str_input_error_magnitude, // 9
                                                        _was_flagged, // 10
                                                        _was_error, // 11
                                                        Environment.NewLine // 12
                                                        ));
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
        CheckCell5          = 0,    // 0.05
        CheckCell10         = 1,    // 0.10
        NormalPerRange      = 2,    //normal analysis of inputs on a per-range granularity
        NormalAllInputs     = 3,    //normal analysis on the entire set of inputs
        CheckCellN          = 4     // CheckCell, n steps
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
        private double _significance;
        private bool _all_outputs;
        private bool _weighted;
        private double _num_max_err_diff_mag;
        private double _str_max_err_diff_mag;
        private double _num_max_output_diff_mag;
        private double _str_max_output_diff_mag;
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

                    double total_error = Utility.CalculateTotalError(correct_outputs, incorrect_outputs);

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

        public void Run(int nboots,                 // number of bootstraps
                        string xlfile,              // name of the workbook
                        double significance,        // significance threshold for test
                        CutoffKind ck,              // kind of threshold function to use
                        Excel.Application app,      // reference to Excel app
                        Classification c,           // data from which to generate errors
                        Random r,                   // a random number generator
                        AnalysisType analysisType,  // the type of analysis to run
                        bool weighted,              // should we weigh things?
                        bool all_outputs,           // if !all_outputs, we only consider terminal outputs
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
            _significance = significance;
            _all_outputs = all_outputs;
            _weighted = weighted;

            //Now we want to inject the errors from _errors
            InjectValues(app, wb, _errors);

            // save function outputs
            CellDict incorrect_outputs = SaveOutputs(terminal_formula_nodes);

            //Time the removal of errors
            Stopwatch sw = new Stopwatch();
            sw.Start();

            // remove errors until none remain; MODIFIES WORKBOOK
            _user = SimulateUser(nboots, significance, ck, data, original_inputs, _errors, correct_outputs, wb, app, analysisType, weighted, all_outputs, max_duration_in_ms, sw, logfile);

            sw.Stop();
            TimeSpan elapsed = sw.Elapsed;
            _analysis_time = elapsed.TotalSeconds;

            // save partially-corrected outputs
            var partially_corrected_outputs = SaveOutputs(terminal_formula_nodes);

            // compute total relative error
            _error = Utility.CalculateNormalizedError(correct_outputs, partially_corrected_outputs, _user.max_errors);
            _total_relative_error = Utility.TotalRelativeError(_error);

            // compute starting total relative error (normalized by max_errors)
            ErrorDict starting_error = Utility.CalculateNormalizedError(correct_outputs, incorrect_outputs, _user.max_errors);
            _initial_total_relative_error = Utility.TotalRelativeError(starting_error);

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

        // For running a simulation from the batch runner
        public void RunFromBatch(int nboots,        // number of bootstraps
                        string xlfile,              // name of the workbook
                        double significance,        // significance threshold for test
                        Excel.Application app,      // reference to Excel app
                        CutoffKind ck,
                        Classification c,           // data from which to generate errors
                        Random r,                   // a random number generator
                        AnalysisType analysisType,  // the type of analysis to run
                        bool weighted,              // should we weigh things?
                        bool all_outputs,           // if !all_outputs, we only consider terminal outputs
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

            // find the error with the largest magnitude
            // this is mostly useful for the single-perturbation experiments
            var num_errs = _errors.Where(pair => Utility.BothNumbers(pair.Value, original_inputs[pair.Key]));
            var str_errs = _errors.Where(pair => !Utility.BothNumbers(pair.Value, original_inputs[pair.Key]));

            _num_max_err_diff_mag = num_errs.Count() != 0 ? num_errs.Select(
                (KeyValuePair<AST.Address, string> pair) =>
                    Utility.NumericalMagnitudeChange(Double.Parse(pair.Value), Double.Parse(original_inputs[pair.Key]))
                    ).Max() : 0;
            _str_max_err_diff_mag = str_errs.Count() != 0 ? str_errs.Select(
                (KeyValuePair<AST.Address, string> pair) =>
                    Utility.StringMagnitudeChange(pair.Value, original_inputs[pair.Key])
                    ).Max() : 0;

            // find the output with the largest magnitude
            var num_outs = correct_outputs.Where(pair => Utility.IsNumber(pair.Value));
            var str_outs = correct_outputs.Where(pair => !Utility.IsNumber(pair.Value));

            _num_max_output_diff_mag = num_outs.Count() != 0 ? num_outs.Select(
                (KeyValuePair<AST.Address, string> pair) =>
                    Utility.NumericalMagnitudeChange(Double.Parse(pair.Value), Double.Parse(correct_outputs[pair.Key]))
                    ).Max() : 0;
            _str_max_output_diff_mag = str_outs.Count() != 0 ? str_outs.Select(
                (KeyValuePair<AST.Address, string> pair) =>
                    Utility.StringMagnitudeChange(pair.Value, correct_outputs[pair.Key])
                    ).Max() : 0;
                
            Run(nboots, xlfile, significance, ck, app, c, r, analysisType, weighted, all_outputs, data, wb, terminal_formula_nodes, terminal_input_nodes, original_inputs, correct_outputs, max_duration_in_ms, logfile);
        }

        public double RemainingError()
        {
            return _total_relative_error / _initial_total_relative_error;
        }

        public static String HeaderRowForCSV()
        {
            return String.Format("{0},{1},{2},{3},{4},{5},{6},{7},{8},{9},{10},{11},{12},{13},{14},{15},{16},{17},{18},{19},{20},{21},{22},{23}{24}",
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
                                 "graph_construct_seconds",                     //14
                                 "analysis_seconds",                            //15
                                 "analysis_type",                               //16
                                 "significance",                                //17
                                 "all_outputs",                                 //18
                                 "weighted",                                    //19
                                 "num_max_err_diff_mag",                        //20
                                 "str_max_err_diff_mag",                        //21
                                 "num_max_out_diff_mag",                        //22
                                 "str_max_out_diff_mag",                        //23
                                 Environment.NewLine                            //24
                                 );                                 
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
                    _significance + "," +
                    _all_outputs + "," +
                    _weighted + "," +
                    _num_max_err_diff_mag + "," +
                    _str_max_err_diff_mag + "," +
                    _num_max_output_diff_mag + "," +
                    _str_max_output_diff_mag +
                    Environment.NewLine;
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

       





 

        



        // remove errors until none remain
        private UserResults SimulateUser(int nboots,
                                         double significance,
                                         CutoffKind ck,
                                         AnalysisData data,
                                         CellDict original_inputs,
                                         CellDict errord,
                                         CellDict correct_outputs,
                                         Excel.Workbook wb,
                                         Excel.Application app,
                                         AnalysisType analysis_type,
                                         bool weighted,
                                         bool all_outputs,
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
            Utility.UpdatePerFunctionMaxError(correct_outputs, incorrect_outputs, max_errors);

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
                if (analysis_type == AnalysisType.CheckCell5 ||
                    analysis_type == AnalysisType.CheckCell10
                    )

                {
                    flagged_cell = SimulationStep.CheckCell_Step(o,
                                                  significance,
                                                  ck,
                                                  nboots,
                                                  data,
                                                  app,
                                                  weighted,
                                                  all_outputs,
                                                  correction_made,
                                                  known_good,
                                                  ref filtered_high_scores,
                                                  max_duration_in_ms,
                                                  sw);
                } else if (analysis_type == AnalysisType.NormalPerRange)
                {
                    flagged_cell = SimulationStep.NormalPerRange_Step(data, wb, known_good, max_duration_in_ms, sw);
                }
                else if (analysis_type == AnalysisType.NormalAllInputs)
                {
                    flagged_cell = SimulationStep.NormalAllOutputs_Step(data, app, wb, known_good, max_duration_in_ms, sw);
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

                        Utility.UpdatePerFunctionMaxError(correct_outputs, partially_corrected_outputs, max_errors);
                        
                        // compute total error after applying this correction
                        var current_total_error = Utility.CalculateTotalError(correct_outputs, partially_corrected_outputs);
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
                    var output_error_magnitude = Utility.MeanErrorMagnitude(partially_corrected_outputs, correct_outputs);
                    // compute input error magnitude
                    double num_input_error_magnitude;
                    double str_input_error_magnitude;
                    if (errord.ContainsKey(flagged_cell))
                    {
                        if (Utility.BothNumbers(errord[flagged_cell], original_inputs[flagged_cell]))
                        {
                            num_input_error_magnitude = Utility.NumericalMagnitudeChange(Double.Parse(errord[flagged_cell]), Double.Parse(original_inputs[flagged_cell]));
                            str_input_error_magnitude = 0;
                        }
                        else
                        {
                            num_input_error_magnitude = 0;
                            str_input_error_magnitude = Utility.StringMagnitudeChange(errord[flagged_cell], original_inputs[flagged_cell]);
                        }
                    }
                    else
                    {
                        num_input_error_magnitude = 0;
                        str_input_error_magnitude = 0;
                    }

                    // write error log
                    var logentry = new LogEntry(analysis_type,
                                                wb.Name,
                                                flagged_cell,
                                                original_inputs[flagged_cell],
                                                errord.ContainsKey(flagged_cell) ? errord[flagged_cell] : original_inputs[flagged_cell],
                                                output_error_magnitude,
                                                num_input_error_magnitude,
                                                str_input_error_magnitude,
                                                true,
                                                correction_made,
                                                significance,
                                                ck.Threshold);
                    logentry.WriteLog(logfile);
                    _error_log.Add(logentry);
                }
            }

            // find all of the false negatives
            o.false_negatives = GetFalseNegatives(o.true_positives, o.false_positives, errord);
            o.max_errors = max_errors;

            var last_out_err_mag = Utility.MeanErrorMagnitude(partially_corrected_outputs, correct_outputs);

            // write out all false negative information
            foreach (AST.Address fn in o.false_negatives)
            {
                double num_input_error_magnitude;
                double str_input_error_magnitude;
                if (Utility.BothNumbers(errord[fn], original_inputs[fn]))
                {
                    num_input_error_magnitude = Utility.NumericalMagnitudeChange(Double.Parse(errord[fn]), Double.Parse(original_inputs[fn]));
                    str_input_error_magnitude = 0;
                }
                else
                {
                    num_input_error_magnitude = 0;
                    str_input_error_magnitude = Utility.StringMagnitudeChange(errord[fn], original_inputs[fn]);
                }

                // write error log
                _error_log.Add(new LogEntry(analysis_type,
                                            wb.Name,
                                            fn,
                                            original_inputs[fn],
                                            errord[fn],
                                            last_out_err_mag,
                                            num_input_error_magnitude,
                                            str_input_error_magnitude,
                                            false,
                                            true,
                                            significance,
                                            ck.Threshold));
            }
            return o;
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

        // save all of the values of the spreadsheet that
        // participate in any computation
        public static CellDict SaveInputs(AnalysisData graph)
        {
            try
            {
                var cd = new CellDict();
                foreach (var node in graph.allComputationCells())
                {
                    cd.Add(node.GetAddress(), node.getCOMValueAsString());
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
                    String fstring = formula_cell.getCOMObject().Formula;
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

                    double total_error = Utility.CalculateTotalError(correct_outputs, incorrect_outputs);

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

        public struct PrepData
        {
            public AnalysisData graph;
            public CellDict original_inputs;
            public CellDict correct_outputs;
            public TreeNode[] terminal_input_nodes;
            public TreeNode[] terminal_formula_nodes;
        }

        public static PrepData PrepSimulation(Excel.Application app, Excel.Workbook wbh, ProgBar pb)
        {
            // build graph
            var graph = DataDebugMethods.ConstructTree.constructTree(wbh, app);
            if (graph.ContainsLoop())
            {
                throw new DataDebugMethods.ContainsLoopException();
            }
            pb.IncrementProgress(16);

            // get terminal input and terminal formula nodes once
            var terminal_input_nodes = graph.TerminalInputNodes();
            var terminal_formula_nodes = graph.TerminalFormulaNodes(true);  ///the boolean indicates whether to use all outputs or not

            if (terminal_input_nodes.Length == 0)
            {
                throw new NoRangeInputs();
            }

            if (terminal_formula_nodes.Length == 0)
            {
                throw new NoFormulas();
            }

            // save original spreadsheet state
            CellDict original_inputs = UserSimulation.Simulation.SaveInputs(graph);

            // force a recalculation before saving outputs, otherwise we may
            // erroneously conclude that the procedure did the wrong thing
            // based solely on Excel floating-point oddities
            UserSimulation.Simulation.InjectValues(app, wbh, original_inputs);

            // save function outputs
            CellDict correct_outputs = UserSimulation.Simulation.SaveOutputs(terminal_formula_nodes);

            return new PrepData() {
                graph = graph,
                original_inputs = original_inputs,
                correct_outputs = correct_outputs,
                terminal_input_nodes = terminal_input_nodes,
                terminal_formula_nodes = terminal_formula_nodes
            };
        }

        public static void RunSimulationPaperMain(Excel.Application app, Excel.Workbook wbh, int nboots, double significance, double threshold, UserSimulation.Classification c, Random r, String outfile, long max_duration_in_ms, String logfile, ProgBar pb)
        {
            // record intitial state of spreadsheet
            var prepdata = PrepSimulation(app, wbh, pb);

            // generate errors
            CellDict errors = UserSimulation.Simulation.GenImportantErrors(prepdata.terminal_formula_nodes,
                                                               prepdata.original_inputs,
                                                               5,
                                                               prepdata.correct_outputs,
                                                               app,
                                                               wbh,
                                                               c);
            // run paper simulations
            RunSimulation(app, wbh, nboots, significance, threshold, c, r, outfile, max_duration_in_ms, logfile, pb, prepdata, errors);
        }

        public static void RunProportionExperiment(Excel.Application app, Excel.Workbook wbh, int nboots, double significance, double threshold, UserSimulation.Classification c, Random r, String outfile, long max_duration_in_ms, String logfile, ProgBar pb)
        {
            // record intitial state of spreadsheet
            var prepdata = PrepSimulation(app, wbh, pb);

            // init error generator
            var eg = new ErrorGenerator();

            // get inputs as an array of addresses to facilitate random selection
            // DATA INPUTS ONLY
            var inputs = prepdata.graph.TerminalInputCells().Select(n => n.GetAddress()).ToArray<AST.Address>();

            // sanity check: all of the inputs should also be in prepdata.original_inputs
            foreach (AST.Address addr in inputs)
            {
                if (!prepdata.original_inputs.ContainsKey(addr))
                {
                    throw new Exception("Missing address!");
                }
            }
            
            for (int i = 0; i < 100; i++)
            {
                // randomly choose an input address
                AST.Address rand_addr = inputs[r.Next(inputs.Length)];

                // get the value
                String input_value = prepdata.original_inputs[rand_addr];

                // perturb it
                String erroneous_input = eg.GenerateErrorString(input_value, c);

                // create an error dictionary with this one perturbed value
                var errors = new CellDict();
                errors.Add(rand_addr, erroneous_input);

                // run simulations; simulation code does insertion of errors and restore of originals
                RunSimulation(app, wbh, nboots, significance, threshold, c, r, outfile, max_duration_in_ms, logfile, pb, prepdata, errors);
            }
        }

        public static void RunSimulation(Excel.Application app, Excel.Workbook wbh, int nboots, double significance, double threshold, UserSimulation.Classification c, Random r, String outfile, long max_duration_in_ms, String logfile, ProgBar pb, PrepData prepdata, CellDict errors)
        {
            pb.IncrementProgress(16);

            // write header if needed
            if (!System.IO.File.Exists(outfile))
            {
                System.IO.File.AppendAllText(outfile, HeaderRowForCSV());
            }

            // CheckCell weighted, all outputs, quantile
            var s_1 = new UserSimulation.Simulation();
            s_1.RunFromBatch(nboots,                                   // number of bootstraps
                                wbh.FullName,                          // Excel filename
                                significance,                          // statistical significance threshold for hypothesis test
                                app,                                   // Excel.Application
                                new QuantileCutoff(0.05),              // max % extreme values to flag
                                c,                                     // classification data
                                r,                                     // random number generator
                                UserSimulation.AnalysisType.CheckCell5,// analysis type
                                true,                                  // weighted analysis
                                true,                                  // use all outputs for analysis
                                prepdata.graph,                                 // AnalysisData
                                wbh,                                   // Excel.Workbook
                                errors,                                // pre-generated errors
                                prepdata.terminal_input_nodes,                  // input range nodes
                                prepdata.terminal_formula_nodes,                // output nodes
                                prepdata.original_inputs,                       // original input values
                                prepdata.correct_outputs,                       // original output values
                                max_duration_in_ms,                    // max duration of simulation 
                                logfile);
            System.IO.File.AppendAllText(outfile, s_1.FormatResultsAsCSV());
            pb.IncrementProgress(16);

            // CheckCell weighted, all outputs, quantile
            var s_4 = new UserSimulation.Simulation();
            s_4.RunFromBatch(nboots,                                   // number of bootstraps
                                wbh.FullName,                          // Excel filename
                                significance,                          // statistical significance of threshold
                                app,                                   // Excel.Application
                                new QuantileCutoff(0.10),              // max % extreme values to flag
                                c,                                     // classification data
                                r,                                     // random number generator
                                UserSimulation.AnalysisType.CheckCell10,// analysis type
                                true,                                  // weighted analysis
                                true,                                  // use all outputs for analysis
                                prepdata.graph,                                 // AnalysisData
                                wbh,                                   // Excel.Workbook
                                errors,                                // pre-generated errors
                                prepdata.terminal_input_nodes,                  // input range nodes
                                prepdata.terminal_formula_nodes,                // output nodes
                                prepdata.original_inputs,                       // original input values
                                prepdata.correct_outputs,                       // original output values
                                max_duration_in_ms,                    // max duration of simulation 
                                logfile);
            System.IO.File.AppendAllText(outfile, s_4.FormatResultsAsCSV());
            pb.IncrementProgress(16);

            // Normal, all inputs
            var s_2 = new UserSimulation.Simulation();
            s_2.RunFromBatch(nboots,                                   // irrelevant
                                wbh.FullName,                              // Excel filename
                                significance,                          // normal cutoff?
                                app,                                   // Excel.Application
                                new NormalCutoff(threshold),           // ??
                                c,                                     // classification data
                                r,                                     // random number generator
                                UserSimulation.AnalysisType.NormalAllInputs,   // analysis type
                                true,                                  // irrelevant
                                true,                                  // irrelevant
                                prepdata.graph,                                 // AnalysisData
                                wbh,                                   // Excel.Workbook
                                errors,                                // pre-generated errors
                                prepdata.terminal_input_nodes,                  // input range nodes
                                prepdata.terminal_formula_nodes,                // output nodes
                                prepdata.original_inputs,                       // original input values
                                prepdata.correct_outputs,                       // original output values
                                max_duration_in_ms,                    // max duration of simulation 
                                logfile);
            System.IO.File.AppendAllText(outfile, s_2.FormatResultsAsCSV());
            pb.IncrementProgress(16);

            // Normal, range inputs
            var s_3 = new UserSimulation.Simulation();
            s_3.RunFromBatch(nboots,                                   // irrelevant
                                wbh.FullName,                              // Excel filename
                                significance,                          // normal cutoff?
                                app,                                   // Excel.Application
                                new NormalCutoff(threshold),           // ??
                                c,                                     // classification data
                                r,                                     // random number generator
                                UserSimulation.AnalysisType.NormalPerRange,   // analysis type
                                true,                                  // irrelevant
                                true,                                  // irrelevant
                                prepdata.graph,                                 // AnalysisData
                                wbh,                                   // Excel.Workbook
                                errors,                                // pre-generated errors
                                prepdata.terminal_input_nodes,                  // input range nodes
                                prepdata.terminal_formula_nodes,                // output nodes
                                prepdata.original_inputs,                       // original input values
                                prepdata.correct_outputs,                       // original output values
                                max_duration_in_ms,                    // max duration of simulation 
                                logfile);
            System.IO.File.AppendAllText(outfile, s_3.FormatResultsAsCSV());
            pb.IncrementProgress(20);
        }
    }
}
