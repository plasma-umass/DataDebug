using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
using System.Text.RegularExpressions;
using Microsoft.FSharp.Core;

namespace DataDebugMethods
{
    public static class Utility
    {
        public static AST.Reference ParseReferenceOfXLRange(Excel.Range rng)
        {
            string rng_r1c1 = rng.Address[true, true, Excel.XlReferenceStyle.xlR1C1, false];
            FSharpOption<AST.Reference> r = ExcelParser.GetReference(rng_r1c1, rng.Worksheet);

            if (FSharpOption<AST.Reference>.get_IsNone(r))
            {
                throw new Exception("Unimplemented address feature in address string: '" + rng_r1c1 + "'");
            }

            return r.Value;
        }

        public static AST.Address ParseXLAddress(Excel.Range rng)
        {
            // we'll get an exception from the parser if this is not, in fact an address
            string rng_r1c1 = rng.Address[true, true, Excel.XlReferenceStyle.xlR1C1, false];
            return ExcelParser.GetAddress(rng_r1c1, rng.Worksheet);
        }

        public static bool InsideRectangle(Excel.Range rng, AST.Reference rect)
        {
            return ParseReferenceOfXLRange(rng).InsideRef(rect);
        }

        public static bool InsideUsedRange(Excel.Range rng)
        {
            return InsideRectangle(rng, UsedRange(rng));
        }

        public static AST.Reference UsedRange(Excel.Range rng)
        {
            return ParseReferenceOfXLRange(rng.Worksheet.UsedRange);
        }
    }
}
