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
        public FunctionOutput<string>[] FastReplace(Excel.Range com, InputSample original, InputSample sample, TreeNode[] outputs, ref int hits)
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
            var sz = com.Cells.Count;
            foreach (Range cell in com)
            {
                if (i >= input.Length())
                {
                    System.Windows.Forms.MessageBox.Show("The parent COM range object has " + sz + " cells, but the saved input sample only has " + input.Length());
                }
                Debug.Assert(i < input.Length());
                cell.Value2 = input.GetInput(i);
                i++;
            }
        }
    }
}
