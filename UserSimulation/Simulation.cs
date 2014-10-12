using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using CellDict = System.Collections.Generic.Dictionary<AST.Address, string>;
using TreeScore = System.Collections.Generic.Dictionary<AST.Address, int>;
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
        public Dictionary<AST.Address, Tuple<string, double>> TopOfKErrors(AST.Address[] terminal_formula_nodes, CellDict inputs, int k, CellDict correct_outputs, Excel.Application app, Excel.Workbook wb, string classification_file, DAG dag)
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
                    Utility.InjectValues(app, wb, cd);

                    // save function outputs
                    CellDict incorrect_outputs = Utility.SaveOutputs(terminal_formula_nodes, dag);

                    //remove the typo that was introduced
                    cd.Clear();
                    cd.Add(addr, orig_value);
                    Utility.InjectValues(app, wb, cd);

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

        // returns the number of cells inspected
        public int Run(int nboots,                 // number of bootstraps
                        string xlfile,              // name of the workbook
                        double significance,        // significance threshold for test
                        CutoffKind ck,              // kind of threshold function to use
                        Excel.Application app,      // reference to Excel app
                        Classification c,           // data from which to generate errors
                        Random r,                   // a random number generator
                        AnalysisType analysisType,  // the type of analysis to run
                        bool weighted,              // should we weigh things?
                        bool all_outputs,           // if !all_outputs, we only consider terminal outputs
                        DAG dag,
                        Excel.Workbook wb,
                        AST.Address[] terminal_formula_cells,
                        AST.Range[] terminal_input_vectors,
                        CellDict original_inputs,
                        CellDict correct_outputs,
                        long max_duration_in_ms,
                        String logfile,              //filename for the output log
                        ProgBar pb
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
            Utility.InjectValues(app, wb, _errors);

            // save function outputs
            CellDict incorrect_outputs = Utility.SaveOutputs(terminal_formula_cells, dag);

            //Time the removal of errors
            Stopwatch sw = new Stopwatch();
            sw.Start();

            // remove errors until none remain; MODIFIES WORKBOOK
            _user = SimulateUser(nboots, significance, ck, dag, original_inputs, _errors, correct_outputs, wb, app, analysisType, weighted, all_outputs, max_duration_in_ms, sw, logfile, pb);

            sw.Stop();
            TimeSpan elapsed = sw.Elapsed;
            _analysis_time = elapsed.TotalSeconds;

            // save partially-corrected outputs
            var partially_corrected_outputs = Utility.SaveOutputs(terminal_formula_cells, dag);

            // compute total relative error
            _error = Utility.CalculateNormalizedError(correct_outputs, partially_corrected_outputs, _user.max_errors);
            _total_relative_error = Utility.TotalRelativeError(_error);

            // compute starting total relative error (normalized by max_errors)
            ErrorDict starting_error = Utility.CalculateNormalizedError(correct_outputs, incorrect_outputs, _user.max_errors);
            _initial_total_relative_error = Utility.TotalRelativeError(starting_error);

            // effort
            _max_effort = dag.allCells().Length;
            _effort = (_user.true_positives.Count + _user.false_positives.Count);
            _expended_effort = (double)_effort / (double)_max_effort;

            // compute average precision
            // AveP = (\sum_{k=1}^n (P(k) * rel(k))) / |total positives|
            // where P(k) is the precision at threshold k,
            // rel(k) = \{ 1 if item at k is a true positive, 0 otherwise
            _average_precision = _user.PrecRel_at_k.Sum() / (double)_errors.Count;

            // restore original values
            Utility.InjectValues(app, wb, original_inputs);

            _tree_construct_time = dag.AnalysisMilliseconds / 1000.0;
            // flag that we're done; safe to print output results
            _simulation_run = true;

            // return the number of cells inspected
            return _effort;
        }

        // For running a simulation from the batch runner
        // returns the number of cells inspected
        public int RunFromBatch(int nboots,        // number of bootstraps
                        string xlfile,              // name of the workbook
                        double significance,        // significance threshold for test
                        Excel.Application app,      // reference to Excel app
                        CutoffKind ck,
                        Classification c,           // data from which to generate errors
                        Random r,                   // a random number generator
                        AnalysisType analysisType,  // the type of analysis to run
                        bool weighted,              // should we weigh things?
                        bool all_outputs,           // if !all_outputs, we only consider terminal outputs
                        DAG dag,          // the computation tree of the spreadsheet
                        Excel.Workbook wb,          // the workbook being analyzed
                        CellDict errors,            // the errors that will be introduced in the spreadsheet
                        AST.Range[] terminal_input_vectors,   // the inputs
                        AST.Address[] terminal_formula_cells, // the outputs
                        CellDict original_inputs,          // original values of the inputs
                        CellDict correct_outputs,          // the correct outputs
                        long max_duration_in_ms,
                        String logfile              //filename for the output log
                       )
        {
            if (terminal_input_vectors.Length == 0)
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

            return Run(nboots, xlfile, significance, ck, app, c, r, analysisType, weighted, all_outputs, dag, wb, terminal_formula_cells, terminal_input_vectors, original_inputs, correct_outputs, max_duration_in_ms, logfile, null);
        }

        public double RemainingError()
        {
            return _total_relative_error / _initial_total_relative_error;
        }

        public static String HeaderRowForCSV()
        {
            return String.Format("{0},{1},{2},{3},{4},{5},{6},{7},{8},{9},{10},{11},{12},{13},{14},{15},{16},{17},{18},{19},{20},{21}{22}",
                                 "workbook_name",                               //0
                                 "initial_total_relative_error",                //1
                                 "total_relative_error",                        //2
                                 "remaining_error",                             //3
                                 "effort",                                      //4
                                 "max_effort",                                  //5
                                 "expended_effort",                             //6
                                 "number_of_errors",                            //7
                                 "true_positives",                              //8
                                 "false_positives",                             //9      
                                 "false_negatives",                             //10
                                 "average_precision",                           //11
                                 "graph_construct_seconds",                     //12
                                 "analysis_seconds",                            //13
                                 "analysis_type",                               //14
                                 "significance",                                //15
                                 "all_outputs",                                 //16
                                 "weighted",                                    //17
                                 "num_max_err_diff_mag",                        //18
                                 "str_max_err_diff_mag",                        //19
                                 "num_max_out_diff_mag",                        //20
                                 "str_max_out_diff_mag",                        //21
                                 Environment.NewLine                            //22
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
                                         DAG dag,
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
                                         String logfile,
                                         ProgBar pb
                                        )
        {
            // init user results data structure
            var o = new UserResults();
            HashSet<AST.Address> known_good = new HashSet<AST.Address>();

            // initialize procedure
            var errors_remain = true;
            var max_errors = new ErrorDict();
            var incorrect_outputs = Utility.SaveOutputs(dag.terminalFormulaNodes(all_outputs), dag);
            var errors_found = 0;
            var number_of_true_errors = errord.Count;
            Utility.UpdatePerFunctionMaxError(correct_outputs, incorrect_outputs, max_errors);

            // the corrected state of the spreadsheet
            CellDict partially_corrected_outputs = correct_outputs.ToDictionary(p => p.Key, p => p.Value);

            // remove errors loop
            var cells_inspected = 0;
            List<KeyValuePair<AST.Address, int>> filtered_high_scores = null;
            bool correction_made = true;
            while (errors_remain)
            {
                Console.Write(".");

                AST.Address flagged_cell = null;

                // choose the appropriate test
                if (analysis_type == AnalysisType.CheckCell5 ||
                    analysis_type == AnalysisType.CheckCell10
                    )

                {
                    flagged_cell = SimulationStep.CheckCell_Step(o,
                                                  significance,
                                                  ck,
                                                  nboots,
                                                  dag,
                                                  app,
                                                  weighted,
                                                  all_outputs,
                                                  correction_made,
                                                  known_good,
                                                  ref filtered_high_scores,
                                                  max_duration_in_ms,
                                                  sw,
                                                  pb);
                } else if (analysis_type == AnalysisType.NormalPerRange)
                {
                    flagged_cell = SimulationStep.NormalPerRange_Step(dag, wb, known_good, max_duration_in_ms, sw);
                }
                else if (analysis_type == AnalysisType.NormalAllInputs)
                {
                    flagged_cell = SimulationStep.NormalAllOutputs_Step(dag, app, wb, known_good, max_duration_in_ms, sw);
                }

                // stop if the test no longer returns anything or if
                // the test is simply done inspecting based on a fixed threshold
                if (flagged_cell == null || (ck.isCountBased && ck.Threshold == cells_inspected))
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
                        partially_corrected_outputs = Utility.SaveOutputs(dag.terminalFormulaNodes(all_outputs), dag);
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
            o.false_negatives = Utility.GetFalseNegatives(o.true_positives, o.false_positives, errord);
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
