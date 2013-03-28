using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
using Range = Microsoft.Office.Interop.Excel.Range;
using System.Diagnostics;

namespace DataDebugMethods
{
    public class BootMemo
    {
        private Dictionary<InputSample, FunctionOutput<string>[]> _d;
        public BootMemo()
        {
            _d = new Dictionary<InputSample, FunctionOutput<string>[]>();
        }
        public FunctionOutput<string>[] FastReplace(Excel.Range com, InputSample original, InputSample sample, TreeNode[] outputs, ref int hits, bool replace_original)
        {
            FunctionOutput<string>[] fo_arr;
            if (!_d.TryGetValue(sample, out fo_arr))
            {
                // replace the COM value
                ReplaceExcelRange(com, sample);

                // initialize array
                fo_arr = new FunctionOutput<string>[outputs.Length];

                // grab all outputs
                for (var k = 0; k < outputs.Length; k++)
                {
                    // save the output
                    fo_arr[k] = new FunctionOutput<string>(outputs[k].getCOMValueAsString(), sample.GetExcludes());
                }

                // Add function values to cache
                _d.Add(sample, fo_arr);

                // restore the COM value
                if (replace_original)
                {
                    ReplaceExcelRange(com, original);
                }
            }
            else
            {
                hits += 1;
            }
            return fo_arr;
        }

        //public static void ReplaceExcelRange(Range com, InputSample input)
        //{
        //    var i = 0;
        //    foreach (Range cell in com)
        //    {
        //        cell.Value2 = input.GetInput(i);
        //        i++;
        //    }
        //}

        public static void ReplaceExcelRange(Range com, InputSample input)
        {
            com.Value2 = input.GetInputArray();

            // if all of the values happen to be numeric, write them as
            // doubles to avoid Excel errors
            //var strarr = input.GetInputArray();
            //var numarr = new double[strarr.GetLength(0),strarr.GetLength(1)];
            //try {
            //    for(var i = 0; i < numarr.GetLength(0); i++)
            //    {
            //        numarr[i,0] = System.Convert.ToDouble(strarr[i,0]);
            //    }
            //    com.Value2 = numarr;
            //} catch {
            //    com.Value2 = strarr;
            //}
        }
    }
}
