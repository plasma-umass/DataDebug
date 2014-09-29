using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
using AddrDict = System.Collections.Generic.Dictionary<Microsoft.Office.Interop.Excel.Range, AST.Address>;

namespace DataDebugMethods
{
    public class AddressCache
    {
        AddrDict _cache = new AddrDict();

        public AddressCache(Excel.Workbook wb, Excel.Application app)
        {
            // get workbook name
            var wbname = new Microsoft.FSharp.Core.FSharpOption<String>(wb.Name);
            var path = new Microsoft.FSharp.Core.FSharpOption<String>(wb.Path);

            // init R1C1 extractor
            var regex = new System.Text.RegularExpressions.Regex("^R([0-9]+)C([0-9]+)$");

            foreach (Excel.Worksheet worksheet in wb.Worksheets)
            {
                // get used range
                Excel.Range rng = worksheet.UsedRange;

                // get dimensions
                var left = rng.Column;                      // 1-based left-hand y coordinate
                var right = rng.Columns.Count + left - 1;   // 1-based right-hand y coordinate
                var top = rng.Row;                          // 1-based top x coordinate
                var bottom = rng.Rows.Count + top - 1;      // 1-based bottom x coordinate

                // get worksheet name
                var wsname = new Microsoft.FSharp.Core.FSharpOption<String>(worksheet.Name);

                // init
                int x_max = right - left;
                int x = -1;
                int y = 0;


                // for each COM object in the used range, create an address object
                // WITHOUT calling any methods on the COM object itself
                foreach (Excel.Range cell in rng)
                {
                    // The basic idea here is that we know how Excel iterates over collections
                    // of cells.  The Excel.Range returned by UsedRange is always rectangular.
                    // Thus we can calculate the addresses of each COM cell reference without
                    // needing to incur the overhead of actually asking it for its address.
                    var x_old = x;
                    x = (x + 1) % (x_max + 1);
                    // increment y if x wrapped
                    y = x < x_old ? y + 1 : y;

                    int c = x + left;
                    int r = y + top;

                    var addr = AST.Address.NewFromR1C1(r, c, wsname, wbname, path);
                    _cache.Add(cell, addr);
                }
            }
        }

        public AST.Address GetAddressOfCell(Excel.Range cell) {
            return _cache[cell];
        }
    }
}
