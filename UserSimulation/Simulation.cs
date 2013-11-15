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

    [Serializable]
    public class Simulation
    {
        private ErrorCondition _exit_state = ErrorCondition.OK;
        private string _exception_message = "";
        private UserResults _user;
        private ErrorDict _error;
        private double _total_relative_error = 0;
        private int _max_effort = 1;
        private int _effort = 0;
        private double _relative_effort = 0;

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
            return _relative_effort;
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
        public Dictionary<TreeNode, Tuple<string, double>> TopOfKErrors(AnalysisData data, int k, CellDict correct_outputs, Excel.Application app, Excel.Workbook wb)
        {
            var eg = new ErrorGenerator();
            var c = Classification.Deserialize();
            var max_error_produced_dictionary = new Dictionary<TreeNode, Tuple<string, double>>();

            foreach (TreeNode inputRange in data.TerminalInputNodes())
            {
                foreach (TreeNode inputNode in inputRange.getInputs())
                {
                    string orig_value = inputNode.getCOMValueAsString();

                    //Load in the classification's dictionaries
                    double max_error_produced = 0.0;
                    string max_error_string = "";

                    // get k strings, in parallel
                    string[] errorstrings = eg.GenerateErrorStrings(orig_value, c, k);

                    for (int i = 0; i < k; i++)
                    {
                        CellDict cd = new CellDict();
                        cd.Add(inputNode.GetAddress(), errorstrings[i]);
                        //inject the typo 
                        InjectValues(app, wb, cd);

                        // save function outputs
                        CellDict incorrect_outputs = SaveOutputs(data.TerminalFormulaNodes(), wb);

                        //remove the typo that was introduced
                        cd.Clear();
                        cd.Add(inputNode.GetAddress(), orig_value);
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
                    max_error_produced_dictionary.Add(inputNode, new Tuple<string, double>(max_error_string, max_error_produced));
                }
            }
            return max_error_produced_dictionary;
        }

        public CellDict GetTopErrors(Dictionary<TreeNode, Tuple<string, double>> max_error_produced_dictionary, double threshold)
        {
            int inputs_count = max_error_produced_dictionary.Count;
            CellDict top_errors = new CellDict();
            while ((top_errors.Count / (double)inputs_count) < threshold)
            {
                double max = 0.0;
                TreeNode max_node = null;
                string max_node_string = "";
                //Find the max_node
                foreach (var kvp in max_error_produced_dictionary)
                {
                    if (kvp.Value.Item2 >= max)
                    {
                        max = kvp.Value.Item2;
                        max_node = kvp.Key;
                        max_node_string = kvp.Value.Item1;
                    }
                }
                max_error_produced_dictionary.Remove(max_node);
                top_errors.Add(max_node.GetAddress(), max_node_string);
            }

            return top_errors;
        }

        // create and run a CheckCell simulation
        public void Run(int nboots, string xlfile, double significance, Excel.Application app, double threshold)
        {
            try
            {
                // open workbook
                Excel.Workbook wb = Utility.OpenWorkbook(xlfile, app);

                // build dependency graph
                var data = ConstructTree.constructTree(app.ActiveWorkbook, app, false);
                if (data.TerminalInputNodes().Length == 0)
                {
                    _exit_state = ErrorCondition.ContainsNoInputs;
                    return;
                }

                // save original spreadsheet state
                CellDict original_inputs = SaveInputs(data.TerminalInputNodes(), wb);
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
                CellDict correct_outputs = SaveOutputs(data.TerminalFormulaNodes(), wb);

                //Look for 'touchy' cells among the inputs:
                //  for each input 
                //      generate K erroneous versions
                //      pick the one that causes the largest total relative error
                //  sort the inputs based on how much total error they are able to produce
                //  pick top 5% for example, and introduce errors
                var max_error_produced_dictionary = TopOfKErrors(data, 10, correct_outputs, app, wb);

                //Now we want to take the inputs that produce the greatest errors
                var top_errors = GetTopErrors(max_error_produced_dictionary, threshold);

                //Now we want to inject the errors in top_errors
                InjectValues(app, wb, top_errors);

                // TODO: save a copy of the workbook for later inspection

                // save function outputs
                CellDict incorrect_outputs = SaveOutputs(data.TerminalFormulaNodes(), wb);

                // remove errors until none remain; MODIFIES WORKBOOK
                _user = SimulateUser(nboots, significance, data, original_inputs, top_errors, correct_outputs, wb, app);

                //// save partially-corrected outputs
                var partially_corrected_outputs = SaveOutputs(data.TerminalFormulaNodes(), wb);

                // compute total relative error
                _error = CalculateNormalizedError(correct_outputs, partially_corrected_outputs, _user.max_errors);
                _total_relative_error = TotalRelativeError(_error);

                // effort
                _max_effort = data.TerminalInputNodes().Length;
                _effort = (_user.true_positives.Count + _user.false_positives.Count);
                _relative_effort = (double)_effort / (double)_max_effort;

                string text_out = wb.Name + "," + _total_relative_error + "," + _effort.ToString() + "," + _max_effort + "," + _relative_effort;
                ToCSV(wb, text_out);

                // close workbook without saving
                wb.Close(false, "", false);
                app.Quit();
                Marshal.ReleaseComObject(wb);
                Marshal.ReleaseComObject(app);
                wb = null;
                app = null;
            }
            catch (Exception e)
            {
                _exit_state = ErrorCondition.Exception;
                _exception_message = e.Message;
            }
        }

        public void ToCSV(Excel.Workbook wb, string out_text)
        {
            string dir_path = wb.Path;
            string file_path = dir_path + "\\Results.csv";
            //if file exists, read it and append to it
            if (System.IO.File.Exists(file_path))
            {
                string text = System.IO.File.ReadAllText(file_path);
                text += "\n" + out_text;
                System.IO.File.WriteAllText(file_path, text);
            }
            //otherwise create the file and write to it
            else
            {
                //System.IO.File.Create(file_path);
                string text = "Workbook name:,Total rel. error:,Effort:,Max effort:,Relative effort:" +
                    "\n" + out_text;
                System.IO.File.WriteAllText(file_path, text);
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

        private static double CalculateTotalError(CellDict correct_outputs, CellDict incorrect_outputs)
        {
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
            public ErrorDict max_errors;
        }

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
                                               Excel.Application app)
        {
            var o = new UserResults();
            o.false_negatives = new HashSet<AST.Address>();
            o.false_positives = new List<AST.Address>();
            o.true_positives = new List<AST.Address>();
            HashSet<AST.Address> known_good = new HashSet<AST.Address>();

            // initialize
            var errors_remain = true;
            var max_errors = new ErrorDict();
            var incorrect_outputs = SaveOutputs(data.TerminalFormulaNodes(), wb);
            UpdatePerFunctionMaxError(correct_outputs, incorrect_outputs, max_errors);

            // remove errors
            while (errors_remain)
            {
                // Get bootstraps
                TreeScore scores = Analysis.Bootstrap(nboots, data, app, true);

                // Compute quantiles based on user-supplied sensitivity
                var quantiles = Analysis.ComputeQuantile<int, TreeNode>(scores.Select(
                    pair => new Tuple<int, TreeNode>(pair.Value, pair.Key))
                );

                // Get top outlier
                var flagged_cell = Analysis.GetTopOutlier(quantiles, known_good, significance);
                if (flagged_cell == null)
                {
                    errors_remain = false;
                }
                else
                {
                    // check to see if the flagged value is actually an error
                    if (errord.ContainsKey(flagged_cell))
                    {
                        o.true_positives.Add(flagged_cell);
                    }
                    else
                    {
                        o.false_positives.Add(flagged_cell);
                    }

                    // correct flagged cell
                    flagged_cell.GetCOMObject(app).Value2 = original_inputs[flagged_cell];
                    var partially_correct_outputs = SaveOutputs(data.TerminalFormulaNodes(), wb);
                    UpdatePerFunctionMaxError(correct_outputs, partially_correct_outputs, max_errors);

                    // mark it as known good
                    known_good.Add(flagged_cell);
                }
            }

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
            var cd = new CellDict();
            foreach (TreeNode input_range in input_ranges)
            {
                foreach (TreeNode cell in input_range.getInputs())
                {
                    // never save formula; there's no point since we don't perturb them
                    var comcell = cell.getCOMObject();
                    if (!comcell.HasFormula)
                    {
                        cd.Add(cell.GetAddress(), cell.getCOMValueAsString());
                    }
                }
            }
            return cd;
        }

        // save spreadsheet outputs to a CellDict
        private static CellDict SaveOutputs(TreeNode[] formula_nodes, Excel.Workbook wb)
        {
            var cd = new CellDict();
            foreach (TreeNode formula_cell in formula_nodes)
            {
                // throw an exception in debug mode, because this should never happen
                Debug.Assert((bool)formula_cell.getCOMObject().HasFormula);
                // save value
                cd.Add(formula_cell.GetAddress(), formula_cell.getCOMValueAsString());
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
