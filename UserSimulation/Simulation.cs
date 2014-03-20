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
    [Serializable]
    public enum ErrorCondition
    {
        OK,
        ContainsNoInputs,
        Exception
    }

    public class SimulationNotRunException : Exception {} 

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
        private string _analysis_type;

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
                // deserialize
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
                    CellDict incorrect_outputs = SaveOutputs(terminal_formula_nodes, wb);

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

        // create and run a CheckCell simulation
        public void Run(int nboots,                 // number of bootstraps
                        string xlfile,              // name of the workbook
                        double significance,        // significance threshold for test
                        Excel.Application app,      // reference to Excel app
                        double threshold,           // percentage of erroneous cells
                        Classification c,           // data from which to generate errors
                        Random r,                   // a random number generator
                        string analysisType,        // the type of analysis to run -- "CheckCell", "Normal", or "Normal2"
                        bool all_outputs            // if !all_outputs, we only consider terminal outputs
                       )
        {
            // set wbname
            _wb_name = xlfile;

            _analysis_type = analysisType;

            // create ErrorGenerator object
            var egen = new ErrorGenerator();

            try
            {
                // open workbook
                Excel.Workbook wb = Utility.OpenWorkbook(xlfile, app);

                // set path
                _wb_path = wb.Path;

                // build dependency graph
                var data = ConstructTree.constructTree(app.ActiveWorkbook, app);
                // get terminal input and terminal formula nodes once
                var terminal_input_nodes = data.TerminalInputNodes();
                var terminal_formula_nodes = data.TerminalFormulaNodes(all_outputs);

                if (terminal_input_nodes.Length == 0)
                {
                    _exit_state = ErrorCondition.ContainsNoInputs;
                    return;
                }

                // save original spreadsheet state
                CellDict original_inputs = SaveInputs(terminal_input_nodes, wb);
                if (original_inputs.Count() == 0)
                {
                    _exit_state = ErrorCondition.ContainsNoInputs;
                    return;
                }

                // force a recalculation before saving outputs, otherwise we may
                // erroneously conclude that the procedure did the wrong thing
                // based solely on Excel floating-point oddities
                InjectValues(app, wb, original_inputs);

                // save function outputs
                CellDict correct_outputs = SaveOutputs(terminal_formula_nodes, wb);

                // generate errors
                _errors = egen.RandomlyGenerateErrors(original_inputs, c, threshold);

                //Now we want to inject the errors from top_errors
                InjectValues(app, wb, _errors);

                // TODO: save a copy of the workbook for later inspection

                // save function outputs
                CellDict incorrect_outputs = SaveOutputs(terminal_formula_nodes, wb);

                // remove errors until none remain; MODIFIES WORKBOOK
                if (analysisType.Equals("CheckCell"))
                {
                    _user = SimulateUser(nboots, significance, data, original_inputs, _errors, correct_outputs, wb, app, "checkcell", false);
                }
                else if (analysisType.Equals("Normal (per range)"))    //Normal (per range)
                {
                    _user = SimulateUser(nboots, significance, data, original_inputs, _errors, correct_outputs, wb, app, "normal", false);
                }
                else
                {
                    _user = SimulateUser(nboots, significance, data, original_inputs, _errors, correct_outputs, wb, app, "normal2", false);
                }

                // save partially-corrected outputs
                var partially_corrected_outputs = SaveOutputs(terminal_formula_nodes, wb);

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
                        if (input_to_range._is_a_cell && !input_to_range.isFormula()) //if this input is a cell and is not a formula, then it is perturbable, so it's in our scope
                        {
                            _cells_in_scope++;
                        }
                    }
                }

                //foreach (var input in data.cell_nodes)
                //{
                //    TreeNode input_node = input.Value;
                //    bool perturbable = false;
                //    foreach (TreeNode output in input_node.getOutputs())
                //    {
                //        //if (terminal_input_nodes.Contains(output))
                //        //if (output.isRange())
                //        //if (output.GetDontPerturb())
                //        if (!output._is_a_cell && !output.GetDontPerturb()) // if the output is a range and it is perturbable
                //        {
                //            perturbable = true;
                //        }
                //    }
                //    if (perturbable == false)
                //    {
                //        _cells_in_scope++;
                //    }
                //}
                //_max_effort = 0;
                //foreach (TreeNode input_range in terminal_input_nodes)
                //{
                //    _max_effort += input_range.getInputs().Count;
                //}
                _effort = (_user.true_positives.Count + _user.false_positives.Count);
                _expended_effort = (double)_effort / (double)_max_effort;

                // compute average precision
                _average_precision = _user.precision_at_step_k.Sum() / (double)_effort;

                // close workbook without saving
                wb.Close(false, "", false);
                Marshal.ReleaseComObject(wb);

                // flag that we're done; safe to print output results
                _simulation_run = true;
            }
            catch (Exception e)
            {
                _exit_state = ErrorCondition.Exception;
                _exception_message = e.Message;
            }
        }

        public double RemainingError()
        {
            return _total_relative_error / _initial_total_relative_error;
        }

        public static String HeaderRowForCSV()
        {
            return String.Format("{0},{1},{2},{3},{4},{5},{6},{7},{8},{9},{10},{11},{12},{13},{14}",
                                 "workbook_name",
                                 "initial_total_relative_error",
                                 "total_relative_error",
                                 "remaining_error",
                                 "effort",
                                 "max_effort",
                                 "cells_in_scope",
                                 "ratio_scope_out_of_total",
                                 "expended_effort",
                                 "number_of_errors",
                                 "true_positives",
                                 "false_positives",
                                 "false_negatives",
                                 "average_precision",
                                 "analysis_type");
        }

        public String FormatResultsAsCSV()
        {
            return _wb_name + "," +                         // workbook name
                    _initial_total_relative_error + "," +   // initial total relative error
                    _total_relative_error + "," +           // final total relative error
                    RemainingError() + "," +                // remaining error
                    _effort.ToString() + "," +              // effort
                    _max_effort + "," +                     // max effort
                    _cells_in_scope + "," +                   // perturbable cells (these are in our scope)
                    (double)_cells_in_scope/(double)_max_effort + "," +                     // proportion of cells that are in scopes of our tool
                    _expended_effort + "," +                // expended effort
                    _errors.Count + "," +                   // number of errors
                    _user.true_positives.Count + "," +      // number of true positives
                    _user.false_positives.Count + "," +     // number of false positives
                    _user.false_negatives.Count + "," +     // number of false negatives
                    _average_precision + "," +              // average precision
                    _analysis_type;                         // anaysis type (CheckCell, Normal per range, normal per worksheet
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
                throw new SimulationNotRunException();
            }
        }

        private static void UpdatePerFunctionMaxError(CellDict correct_outputs, CellDict incorrect_outputs, ErrorDict max_errors)
        {
            foreach (var kvp in correct_outputs)
            {
                var addr = kvp.Key;
                var correct_value = correct_outputs[addr];
                var incorrect_value = incorrect_outputs[addr];
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
        private struct UserResults
        {
            public List<AST.Address> true_positives;
            public List<AST.Address> false_positives;
            public HashSet<AST.Address> false_negatives;
            public ErrorDict max_errors; //Keeps track of the largest errors we observe during the simulation for each output
            public List<double> current_total_error;
            public List<double> precision_at_step_k;
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

        // remove errors until none remain
        private static UserResults SimulateUser(int nboots,
                                               double significance,
                                               AnalysisData data,
                                               CellDict original_inputs,
                                               CellDict errord,
                                               CellDict correct_outputs,
                                               Excel.Workbook wb,
                                               Excel.Application app,
                                               string analysis_type,
                                               bool all_outputs
                                            )
        {
            // init user results data structure
            var o = new UserResults();
            o.false_negatives = new HashSet<AST.Address>();
            o.false_positives = new List<AST.Address>();
            o.true_positives = new List<AST.Address>();
            o.current_total_error = new List<double>();
            o.precision_at_step_k = new List<double>();
            HashSet<AST.Address> known_good = new HashSet<AST.Address>();

            // initialize procedure
            var errors_remain = true;
            var max_errors = new ErrorDict();
            var incorrect_outputs = SaveOutputs(data.TerminalFormulaNodes(all_outputs), wb);
            var errors_found = 0;
            var number_of_true_errors = errord.Count;
            UpdatePerFunctionMaxError(correct_outputs, incorrect_outputs, max_errors);

            // remove errors
            var cells_inspected = 0;
            while (errors_remain)
            {
                cells_inspected += 1;
                Console.Write(".");

                AST.Address flagged_cell = null;

                if (analysis_type.Equals("checkcell"))
                {
                    // Get bootstraps
                    TreeScore scores = Analysis.Bootstrap(nboots, data, app, true, false);
                    /*
                    // Compute quantiles based on user-supplied sensitivity
                    var quantiles = Analysis.ComputeQuantile<int, TreeNode>(scores.Select(
                        pair => new Tuple<int, TreeNode>(pair.Value, pair.Key))
                    );

                    // Get top outlier
                    flagged_cell = Analysis.GetTopOutlier(quantiles, known_good, significance);
                     */
                    var scores_list = scores.OrderByDescending(pair => pair.Value).ToList(); //pair => pair.Key, pair => pair.Value);

                    int start_ptr = 0;
                    int end_ptr = 0;

                    List<KeyValuePair<TreeNode, int>> high_scores = new List<KeyValuePair<TreeNode, int>>();

                    while ((double)start_ptr / scores_list.Count < 1.0 - significance) //the start of this score region is before the cutoff
                    {
                        //while the scores at the start and end pointers are the same, bump the end pointer
                        while (end_ptr < scores_list.Count && scores_list[start_ptr].Value == scores_list[end_ptr].Value)
                        {
                            end_ptr++;
                        }
                        //Now the end_pointer points to the first index with a lower score
                        //If the end pointer is still above the significance cutoff, add all values of this score to the high_scores list
                        if ((double)end_ptr / scores_list.Count < 1.0 - significance)
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
                    var filtered_scores = high_scores.Where(kvp => !known_good.Contains(kvp.Key.GetAddress())).ToList();

                    //AST.Address flagged_cell;
                    if (filtered_scores.Count() != 0)
                    {
                        // get TreeNode corresponding to most unusual score
                        flagged_cell = filtered_scores[0].Key.GetAddress();
                    }
                    else
                    {
                        flagged_cell = null;
                    }
                }
                else if (analysis_type.Equals("normal"))
                {
                    //Generate normal distributions for every input range
                    foreach (var range in data.input_ranges.Values)
                    {
                        var normal_dist = new DataDebugMethods.NormalDistribution(range.getCOMObject());

                        // Get top outlier
                        if (normal_dist.errorsCount() > 0)
                        {
                            for (int i = 0; i < normal_dist.errorsCount(); i++)
                            {
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
                    }
                }
                else if (analysis_type.Equals("normal2"))
                {
                    //Generate normal distributions for every worksheet
                    foreach (Excel.Worksheet ws in wb.Worksheets)
                    {
                        var normal_dist = new DataDebugMethods.NormalDistribution(ws.UsedRange);

                        // Get top outlier
                        if (normal_dist.errorsCount() > 0)
                        {
                            for (int i = 0; i < normal_dist.errorsCount(); i++)
                            {
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
                    }
                }

                if (flagged_cell == null)
                {
                    errors_remain = false;
                }
                else
                {
                    // check to see if the flagged value is actually an error
                    if (errord.ContainsKey(flagged_cell))
                    {
                        errors_found += 1;
                        o.precision_at_step_k.Add(errors_found / (double)cells_inspected);
                        o.true_positives.Add(flagged_cell);

                        // correct flagged cell -- only need to do this if the flagged cell was an error
                        flagged_cell.GetCOMObject(app).Value2 = original_inputs[flagged_cell];
                        var partially_corrected_outputs = SaveOutputs(data.TerminalFormulaNodes(all_outputs), wb);
                        UpdatePerFunctionMaxError(correct_outputs, partially_corrected_outputs, max_errors);
                        
                        // compute total error after applying this correction
                        var current_total_error = CalculateTotalError(correct_outputs, partially_corrected_outputs);
                        o.current_total_error.Add(current_total_error);
                    }
                    else
                    {
                        o.precision_at_step_k.Add(0);
                        o.false_positives.Add(flagged_cell);
                    }

                    // mark it as known good -- at this point the cell has been 
                    //      'inspected' regardless of whether it was an error
                    known_good.Add(flagged_cell);

                    // write CSV line
                    //ToTimeseriesCSV(wb, current_total_error, cells_inspected);
                }
            }

            //Console.Write("\n");

            // find all of the false negatives
            o.false_negatives = GetFalseNegatives(o.true_positives, o.false_positives, errord);
            o.max_errors = max_errors;

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
        private static CellDict SaveInputs(TreeNode[] input_ranges, Excel.Workbook wb)
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
        private static CellDict SaveOutputs(TreeNode[] formula_nodes, Excel.Workbook wb)
        {
            var cd = new CellDict();
            foreach (TreeNode formula_cell in formula_nodes)
            {
                // throw an exception in debug mode, because this should never happen
                Debug.Assert((bool)formula_cell.getCOMObject().HasFormula);

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
        private static void InjectValues(Excel.Application app, Excel.Workbook wb, CellDict values)
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
