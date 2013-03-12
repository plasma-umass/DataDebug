using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
using Range = Microsoft.Office.Interop.Excel.Range;

namespace DataDebugMethods
{
    public class BootMemo
    {
        private Dictionary<InputSample, FunctionOutput[]> _d;
        public BootMemo()
        {
            _d = new Dictionary<InputSample, FunctionOutput[]>();
        }
        public FunctionOutput[] FastReplace(Excel.Range com, InputSample original, InputSample sample, TreeNode[] outputs, ref int hits)
        {
            FunctionOutput[] fo_arr;
            if (!_d.TryGetValue(sample, out fo_arr))
            {
                // replace the COM value
                ReplaceExcelRange(com, sample);

                // initialize array
                fo_arr = new FunctionOutput[outputs.Length];

                // grab all outputs
                for (var k = 0; k < outputs.Length; k++)
                {
                    // save the output
                    fo_arr[k] = new FunctionOutput(outputs[k].getCOMValueAsString(), sample.GetExcludes());
                }

                // Add function values to cache
                _d.Add(sample, fo_arr);

                // restore the COM value
                ReplaceExcelRange(com, original);
            }
            else
            {
                hits += 1;
            }
            return fo_arr;
        }
        private static void ReplaceExcelRange(Range com, InputSample input)
        {
            var i = 0;
            foreach (Range cell in com)
            {
                cell.Value2 = input.GetInput(i);
                i++;
            }
        }
    }
}
