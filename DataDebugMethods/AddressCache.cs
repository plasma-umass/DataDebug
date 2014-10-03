using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using Excel = Microsoft.Office.Interop.Excel;
using AddrDict = System.Collections.Generic.Dictionary<AST.COMRef, AST.Address>;
using RangeDict = System.Collections.Generic.Dictionary<AST.COMRef, AST.Range>;
using COMCellDict = System.Collections.Generic.Dictionary<AST.Address, AST.COMRef>;
using COMRangeDict = System.Collections.Generic.Dictionary<AST.Range, AST.COMRef>;
using FormulaDict = System.Collections.Generic.Dictionary<AST.Address, string>;

namespace DataDebugMethods
{
    public class AddressCache
    {
        private Excel.Application _app;
        private AddrDict _addr_cache = new AddrDict();
        private RangeDict _range_cache = new RangeDict();
        private COMCellDict _com_cell_cache = new COMCellDict();
        private COMRangeDict _com_range_cache = new COMRangeDict();
        private FormulaDict _formulas = new FormulaDict();

        public AddressCache(Excel.Workbook wb, Excel.Application app)
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
                    _addr_cache.Add(cr, addr);
                    _com_cell_cache.Add(addr, cr);
                }
            }
        }

        public AST.COMRef GetCOMObjectForAddress(AST.Address addr)
        {
            return _com_cell_cache[addr];
        }

        public AST.Address[] GetFormulaAddrs()
        {
            return _formulas.Keys.ToArray();
        }

        public COMCellDict GetFormulaDictionary()
        {
            return GetFormulaAddrs().ToDictionary(
                addr => addr,
                addr => _com_cell_cache[addr]
            );
        }

        public AST.COMRef MakeCOMRef(AST.Range rng, AST.COMRef parent)
        {
            // check for the range in the dictionary
            AST.COMRef c;
            if (!_com_range_cache.TryGetValue(rng, out c))
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
                _com_range_cache.Add(rng, c);
                _range_cache.Add(c, rng);
            }
            return c;
        }

        //public static void CreateCellNodesFromRange(TreeNode input_range, TreeNode formula, TreeDict formula_nodes, TreeDict cell_nodes, Excel.Workbook wb, bool ignore_parse_errors)
        //{
        //    foreach (Excel.Range cell in input_range.getCOMObject())
        //    {
        //        var addr = AST.Address.AddressFromCOMObject(cell, formula.getWorkbookObject());

        //        // cell might either be another formula or just a simple data cell;
        //        var d = cell.HasFormula ? formula_nodes : cell_nodes;

        //        // add to appropriate dictionary
        //        TreeNode cell_node;
        //        if (!d.TryGetValue(addr, out cell_node))
        //        {
        //            cell_node = new TreeNode(cell, cell.Worksheet, formula.getWorkbookObject());
        //            d.Add(addr, cell_node);
        //        }

        //        // Allow perturbation of every input_range that contains at least one value
        //        // TODO: fix; the Workbook reference here is not correct in the case of cross-workbook reference;
        //        // that said, having the wrong workbook doesn't actually have any bearing on the correctness of this call
        //        if ((cell.HasFormula && ExcelParserUtility.GetSCFormulaNames((string)cell.Formula, wb.FullName, cell.Worksheet, wb, ignore_parse_errors).Count() > 0)) //|| cell.Value2 != null)
        //        {
        //            input_range.SetDoNotPerturb();
        //        }

        //        // link cell, range, and formula inputs and outputs together
        //        input_range.addInput(cell_node);
        //        cell_node.addOutput(formula);
        //        formula.addInput(cell_node);
        //    }
        //}

        public void CreateCellRefsFromRange(AST.COMRef rng_ref)
        {
            // generate input range addresses so that we do
            // not need to ask COM for them
            var com_rng = _range_cache[rng_ref];
            var input_addrs = com_rng.Addresses();
            foreach (AST.Address addr in input_addrs)
            {
                // look in cache before we try to create anything
                AST.COMRef cell_ref;
                if (!_com_cell_cache.TryGetValue(addr, out cell_ref))
                {
                    var com_cell = addr.GetCOMObject(_app);
                    Excel.Worksheet ws = rng_ref.Worksheet;
                    Excel.Workbook wb = rng_ref.Workbook;
                    string wsname = rng_ref.WorksheetName;
                    string wbname = rng_ref.WorkbookName;
                    string path = rng_ref.Path;
                    string fstr;
                    Microsoft.FSharp.Core.FSharpOption<string> formula;
                    if (_formulas.TryGetValue(addr, out fstr) {
                        formula = new Microsoft.FSharp.Core.FSharpOption<string>(fstr);
                    } else {
                        formula = Microsoft.FSharp.Core.FSharpOption<string>.None;
                    }

                    cell_ref = new AST.COMRef(addr.A1FullyQualified(), wb, ws, com_cell, path, wbname, wsname, formula, 1, 1);

                    // add to caches

                }
            }
        }
    }
}
