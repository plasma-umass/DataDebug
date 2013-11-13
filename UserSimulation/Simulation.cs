﻿using System;
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

        // create and run a CheckCell simulation
        public void Run(int nboots, string filename, double significance, ErrorDB errors, Excel.Application app)
        {
            var errord = ErrorDBToCellDict(errors);

            try
            {
                // open workbook
                Excel.Workbook wb = Utility.OpenWorkbook(filename, app);

                // build dependency graph
                var data = ConstructTree.constructTree(app.ActiveWorkbook, app, true);
                if (data.TerminalInputNodes().Length == 0)
                {
                    _exit_state = ErrorCondition.ContainsNoInputs;
                    return;
                }

                // save original spreadsheet state
                CellDict original_inputs = SaveInputs(data.TerminalInputNodes(), wb);

                // force a recalculation before saving outputs, otherwise we may
                // erroneously conclude that the procedure did the wrong thing
                // based solely on Excel floating-point oddities
                InjectValues(app, wb, original_inputs);

                // save function outputs
                CellDict correct_outputs = SaveOutputs(data.TerminalFormulaNodes(), wb);

                //Look for 'touchy' cells among the inputs:
                //  for each input 
                //      generate K erroneous versions
                //      pick the one which causes the largest total relative error
                //  sort the inputs based on how much total error they are able to produce
                //  pick top 5% for example, and introduce errors
                Dictionary<TreeNode, double> max_error_produced_dictionary = new Dictionary<TreeNode, double>();

                foreach (TreeNode inputNode in data.TerminalInputNodes())
                {
                    string orig_value = inputNode.getCOMValueAsString();
                    var eg = new ErrorGenerator();

                    //Load in the classification's dictionaries
                    var classification = Classification.Deserialize();
                    double max_error_produced = 0.0;
                    for (int i = 0; i < 10; i++)
                    {
                        //Generate error string from orig_value
                        var result = eg.GenerateErrorString(orig_value, classification);
                        //If it's no different from the original, try again
                        if (result.Item1.Equals(orig_value))
                        {
                            i--;
                        }
                        //If there was an error, find the total error in the outputs introduced by it
                        else
                        {
                            // save function outputs
                            CellDict perturbed_outputs = SaveOutputs(data.TerminalFormulaNodes(), wb);
                            // error
                            ErrorDict err_dict = CalculateError(app, correct_outputs, perturbed_outputs);
                            double total_rel_err = TotalRelativeError(err_dict);
                            //keep track of the largest observed max error
                            if (Math.Abs(total_rel_err) > Math.Abs(max_error_produced))
                            {
                                max_error_produced = total_rel_err;
                            }
                        }
                    }
                    //Add entry for this TreeNode in our dictionary with its max_error_produced
                    max_error_produced_dictionary.Add(inputNode, max_error_produced);
                }

                //Sort the dictionary to find the most important inputs
                //max_error_produced_dictionary.OrderBy
                //List top_inputs = [];
                //while top_inputs.count / inputs.count < 5%
                //  s = largest entry in dictionary
                //  dict.remove(s)
                //  top_inputs.add(s)

                //Now we want to inject errors in the top_inputs

                // inject errors
                InjectValues(app, wb, errord);

                // save function outputs
                CellDict incorrect_outputs = SaveOutputs(data.TerminalFormulaNodes(), wb);

                // remove errors until none remain; MODIFIES WORKBOOK
                _user = SimulateUser(nboots, significance, data, original_inputs, errord, app);

                // error
                _error = CalculateError(app, correct_outputs, incorrect_outputs);
                _total_relative_error = TotalRelativeError(_error);

                // effort
                _max_effort = data.TerminalInputNodes().Length;
                _effort = (_user.true_positives.Count + _user.false_positives.Count);
                _relative_effort = (double)_effort / (double)_max_effort;

                // close workbook without saving
                wb.Close(false, "", false);
            }
            catch
            {
                _exit_state = ErrorCondition.Exception;
            }
        }

        [Serializable]
        private struct UserResults
        {
            public List<AST.Address> true_positives;
            public List<AST.Address> false_positives;
            public HashSet<AST.Address> false_negatives;
        }

        private static double TotalRelativeError(ErrorDict error)
        {
            return error.Select(pair => pair.Value).Sum() / (double)error.Count();
        }

        private static ErrorDict CalculateError(Excel.Application app, CellDict correct_outputs, CellDict incorrect_outputs)
        {
            var ed = new ErrorDict();

            foreach (KeyValuePair<AST.Address, string> orig in correct_outputs)
            {
                var addr = orig.Key;
                string original_value = orig.Value;
                string perturbed_value = System.Convert.ToString(addr.GetCOMObject(app).Value2);
                string corrected_value = System.Convert.ToString(correct_outputs[addr]);
                // if the function produces numeric outputs, calculate distance
                if (ExcelParser.isNumeric(original_value) &&
                    ExcelParser.isNumeric(perturbed_value) &&
                    ExcelParser.isNumeric(corrected_value))
                {
                    ed.Add(addr, RelativeNumericError(System.Convert.ToDouble(original_value),
                                                      System.Convert.ToDouble(perturbed_value),
                                                      System.Convert.ToDouble(corrected_value)));
                }
                // calculate indicator function
                else
                {
                    ed.Add(addr, RelativeCategoricalError(original_value, corrected_value));
                }
            }

            return ed;
        }

        // compares the corrected function output against the incorrected output
        // 0 means that the error has been completely corrected; 1 means that
        // the error totally remains
        private static double RelativeNumericError(double original_value, double perturbed_value, double corrected_value)
        {
            //|f(I'') - f(I)| / |f(I') - f(I)|
            return Math.Abs(corrected_value - original_value) / Math.Abs(perturbed_value - original_value);
        }

        private static double RelativeCategoricalError(string original_value, string corrected_value)
        {
            if (String.Equals(original_value, corrected_value))
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
                                               Excel.Application app)
        {
            var o = new UserResults();
            HashSet<AST.Address> known_good = new HashSet<AST.Address>();

            var errors_remain = true;
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

                    // mark it as known good
                    known_good.Add(flagged_cell);
                }
            }

            // find all of the false negatives
            o.false_negatives = GetFalseNegatives(o.true_positives, o.false_positives, errord);

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
                foreach (TreeNode cell in input_range.getChildren())
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
                Debug.Assert(formula_cell.getCOMObject().HasFormula);
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
                    addr.GetCOMObject(app).Value2 = errorstr;
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
