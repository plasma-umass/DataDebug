using System;
using System.Collections.Generic;
using System.Collections.Concurrent;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
using Range = Microsoft.Office.Interop.Excel.Range;
using System.Diagnostics;

namespace DataDebugMethods
{
    public class BootMemo
    {
        private Dictionary<InputSample, FunctionOutput<string>[]> _d = new Dictionary<InputSample, FunctionOutput<string>[]>();
        
        public FunctionOutput<string>[] FastReplace(Excel.Range com, DAG dag, InputSample original, InputSample sample, AST.Address[] outputs, bool replace_original)
        {
            FunctionOutput<string>[] fo_arr;
            if (!_d.TryGetValue(sample, out fo_arr))
            {
                // replace the COM value
                ReplaceExcelRange(com, sample);

                // initialize array
                fo_arr = new FunctionOutput<string>[outputs.Length];

                // grab all outputs
                var fos = DAG.fastOutputRead(com.Worksheet.Parent, dag.getAllFormulaAddrsAsHashSet());
                for (var k = 0; k < outputs.Length; k++)
                {
                    fo_arr[k] = new FunctionOutput<string>(fos[outputs[k]], sample.GetExcludes());
                }

                // Add function values to cache
                // Don't care about return value
                _d.Add(sample, fo_arr);

                // restore the COM value
                if (replace_original)
                {
                    ReplaceExcelRange(com, original);
                }
            }
            return fo_arr;
        }

        public static void ReplaceExcelRange(Range com, InputSample input)
        {
            bool done = false;
            while (!done)
            {
                try
                {
                    com.Value2 = input.GetInputArray();
                    done = true;
                }
                catch (Exception)
                {

                }
            }
        }
    }
}
