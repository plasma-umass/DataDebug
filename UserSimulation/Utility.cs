using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
using TreeNode = DataDebugMethods.TreeNode;
using CellDict = System.Collections.Generic.Dictionary<AST.Address, string>;
using TreeScore = System.Collections.Generic.Dictionary<DataDebugMethods.TreeNode, int>;
using ErrorDict = System.Collections.Generic.Dictionary<AST.Address, double>;

namespace UserSimulation
{
    public static class Utility
    {
        public static CellDict ErrorDBToCellDict(ErrorDB errors)
        {
            var d = new CellDict();
            foreach (Error e in errors.Errors)
            {
                d.Add(e.GetAddress(), e.value);
            }
            return d;
        }

        public static void UpdatePerFunctionMaxError(CellDict correct_outputs, CellDict incorrect_outputs, ErrorDict max_errors)
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
        public static double CalculateTotalError(CellDict correct_outputs, CellDict incorrect_outputs)
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

        //Computes total relative error
        //Each entry in the dictionary is normalized to its max value, so they are all <= 1.0.
        //We sum them up and divide by the total number of entries to get the total relative error
        public static double TotalRelativeError(ErrorDict error)
        {
            return error.Select(pair => pair.Value).Sum() / (double)error.Count();
        }

        public static ErrorDict CalculateNormalizedError(CellDict correct_outputs, CellDict partially_corrected_outputs, ErrorDict max_errors)
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
        public static double RelativeNumericError(double correct_value, double partially_corrected_value, double max_error)
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

        public static double RelativeCategoricalError(string original_value, string partially_corrected_value)
        {
            if (String.Equals(original_value, partially_corrected_value))
            {
                return 0;
            }
            else
            {
                return 1;
            }
        }

        public static bool IsNumber(string value)
        {
            double v;
            return Double.TryParse(value, out v);
        }

        public static bool BothNumbers(string value1, string value2)
        {
            return IsNumber(value1) && IsNumber(value2);
        }

        // this represents the magnitude of the change; much more meaningful
        public static double NumericalMagnitudeChange(double error, double correct)
        {
            if (error - correct != 0)
            {
                // we add a tiny value to c to avoid divide-by-zero
                return Math.Log10(Math.Abs(error - correct) / Math.Abs(correct + 0.000000000001));
            }
            else
            {
                return 0;
            }
        }

        public static double StringMagnitudeChange(string error, string correct)
        {
            return error == correct ? 0 : 1;
        }


        public static double MeanErrorMagnitude(CellDict partially_corrected_outputs, CellDict original_outputs)
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
    }

}
