using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using Excel = Microsoft.Office.Interop.Excel;
using CellRefDict = DataDebugMethods.BiDictionary<AST.Address, AST.COMRef>;
using VectorRefDict = DataDebugMethods.BiDictionary<AST.Range, AST.COMRef>;
using FormulaDict = System.Collections.Generic.Dictionary<AST.Address, string>;
using FormulaInputVectDict = DataDebugMethods.BiDictionary<AST.Address, AST.Range>;
using FormulaInputCellDict = DataDebugMethods.BiDictionary<AST.Address, AST.Address>;

namespace DataDebugMethods
{
    public class DAG
    {
        private Excel.Application _app;
        private CellRefDict _all_cells = new CellRefDict();
        private VectorRefDict _all_vectors = new VectorRefDict();
        private FormulaDict _formulas = new FormulaDict();
        private FormulaInputVectDict _formula_vectors = new FormulaInputVectDict();
        private FormulaInputCellDict _formula_cells = new FormulaInputCellDict();

        public DAG(Excel.Workbook wb, Excel.Application app)
        {
            // get names
            _app = app;
            var wbfullname = wb.FullName;
            var wbname = wb.Name;
            var path = wb.Path;
            var wbname_opt = new Microsoft.FSharp.Core.FSharpOption<String>(wbname);
            var path_opt = new Microsoft.FSharp.Core.FSharpOption<String>(path);

            // init R1C1 extractor
            var regex = new Regex("^R([0-9]+)C([0-9]+)$");

            // init formula validator
            var fn_filter = new Regex("^=", RegexOptions.Compiled);

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
                var wsname = worksheet.Name;
                var wsname_opt = new Microsoft.FSharp.Core.FSharpOption<String>(wsname);

                // init
                int x_max = right - left;
                int y_max = bottom - top;
                int x = -1;
                int y = 0;

                // array read of formula cells
                // note that this is a 1-based 2D multiarray
                object[,] formulas = rng.Formula;

                // for every cell that is actually a formula, add to 
                // formula dictionary
                for (int c = 1; c <= x_max; c++)
                {
                    for (int r = 1; r <= y_max; r++)
                    {
                        var f = (string)formulas[c,r];
                        if (fn_filter.IsMatch(f)) {
                            var addr = AST.Address.NewFromR1C1(r, c, wsname_opt, wbname_opt, path_opt);
                            _formulas.Add(addr, f);
                        }
                    }
                }

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

                    var addr = AST.Address.NewFromR1C1(r, c, wsname_opt, wbname_opt, path_opt);
                    var formula = _formulas.ContainsKey(addr) ? new Microsoft.FSharp.Core.FSharpOption<string>(_formulas[addr]) : Microsoft.FSharp.Core.FSharpOption<string>.None;
                    var cr = new AST.COMRef(addr.A1FullyQualified(), wb, worksheet, cell, path, wbname, wsname, formula, 1, 1);
                    _all_cells.Add(addr, cr);
                }
            }
        }

        public AST.COMRef GetCOMRefForAddress(AST.Address addr)
        {
            return _all_cells[addr];
        }

        public AST.Address[] GetFormulaAddrs()
        {
            return _formulas.Keys.ToArray();
        }

        public AST.COMRef MakeInputVectorCOMRef(AST.Range rng)
        {
            // check for the range in the dictionary
            AST.COMRef c;
            if (!_all_vectors.TryGetValue(rng, out c))
            {
                // otherwise, create and cache it
                Excel.Range com = rng.GetCOMObject(_app);
                Excel.Worksheet ws = com.Worksheet;
                Excel.Workbook wb = ws.Parent;
                string wsname = ws.Name;
                string wbname = wb.Name;
                string path = wb.Path;
                int width = com.Columns.Count;
                int height = com.Rows.Count;

                c = new AST.COMRef(rng.getUniqueID(), wb, ws, com, path, wbname, wsname, Microsoft.FSharp.Core.FSharpOption<string>.None, width, height);
                _all_vectors.Add(rng, c);
            }
            return c;
        }

        public void LinkInputVector(AST.Address formula_addr, AST.Range vector_rng) {
            _formula_vectors.Add(formula_addr, vector_rng);
        }

        public void LinkCell(AST.Address formula_addr, AST.Address input_addr)
        {
            _formula_cells.Add(formula_addr, input_addr);
        }
    }
}
