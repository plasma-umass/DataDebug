using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
using System.Text.RegularExpressions;
using Microsoft.FSharp.Core;

using Workbook = Microsoft.Office.Interop.Excel.Workbook;
using Worksheet = Microsoft.Office.Interop.Excel.Worksheet;

namespace DataDebugMethods
{
    public static class Utility
    {
        public static AST.Reference ParseReferenceOfXLRange(Excel.Range rng, Workbook wb)
        {
            string rng_r1c1 = rng.Address[true, true, Excel.XlReferenceStyle.xlR1C1, false];
            FSharpOption<AST.Reference> r = ExcelParser.GetReference(rng_r1c1, wb, rng.Worksheet);

            if (FSharpOption<AST.Reference>.get_IsNone(r))
            {
                throw new Exception("Unimplemented address feature in address string: '" + rng_r1c1 + "'");
            }

            return r.Value;
        }

        public static AST.Address ParseXLAddress(Excel.Range rng, Workbook wb)
        {
            // we'll get an exception from the parser if this is not, in fact an address
            string rng_r1c1 = rng.Address[true, true, Excel.XlReferenceStyle.xlR1C1, false];
            return ExcelParser.GetAddress(rng_r1c1, wb, rng.Worksheet);
        }

        public static bool InsideRectangle(Excel.Range rng, AST.Reference rect, Workbook wb)
        {
            return ParseReferenceOfXLRange(rng, wb).InsideRef(rect);
        }

        public static bool InsideUsedRange(Excel.Range rng, Workbook wb)
        {
            return InsideRectangle(rng, UsedRange(rng, wb), wb);
        }

        public static AST.Reference UsedRange(Excel.Range rng, Workbook wb)
        {
            return ParseReferenceOfXLRange(rng.Worksheet.UsedRange, wb);
        }
    }
}
