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
        public static AST.Range AddressOfXLRange(Excel.Range rng)
        {
            string rng_r1c1 = rng.Address[true, true, Excel.XlReferenceStyle.xlR1C1, false];
            FSharpOption<AST.Range> r = ExcelParser.GetRange(rng_r1c1);

            if (FSharpOption<AST.Range>.get_IsNone(r))
            {
                throw new Exception("Unimplemented address feature in address string: '" + rng_r1c1 + "'");
            }

            return r.Value;
        }

        public static bool InsideRectangle(Excel.Range rng, AST.Range rect)
        {
            AST.Range addr_rng = AddressOfXLRange(rng);

            bool is_bad = (addr_rng.getXLeft() < rect.getXLeft() ||
                           addr_rng.getYTop() < rect.getYTop() ||
                           addr_rng.getXRight() > rect.getXRight() ||
                           addr_rng.getYBottom() > rect.getYBottom());

            return !is_bad;
        }

        public static bool InsideUsedRange(Excel.Range rng)
        {
            return InsideRectangle(rng, UsedRange(rng));
        }

        public static AST.Range UsedRange(Excel.Range rng)
        {
            return AddressOfXLRange(rng.Worksheet.UsedRange);
        }
    }
}
