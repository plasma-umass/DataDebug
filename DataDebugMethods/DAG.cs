using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using Excel = Microsoft.Office.Interop.Excel;
using CellRefDict = DataDebugMethods.BiDictionary<AST.Address, AST.COMRef>;
using VectorRefDict = DataDebugMethods.BiDictionary<AST.Range, AST.COMRef>;
using FormulaDict = System.Collections.Generic.Dictionary<AST.Address, string>;
using Formula2VectDict = System.Collections.Generic.Dictionary<AST.Address, System.Collections.Generic.HashSet<AST.Range>>;
using Vect2FormulaDict = System.Collections.Generic.Dictionary<AST.Range, System.Collections.Generic.HashSet<AST.Address>>;
using Vect2InputCellDict = System.Collections.Generic.Dictionary<AST.Range, System.Collections.Generic.HashSet<AST.Address>>;
using InputCell2VectDict = System.Collections.Generic.Dictionary<AST.Address, System.Collections.Generic.HashSet<AST.Range>>;
using Formula2InputCellDict = System.Collections.Generic.Dictionary<AST.Address, System.Collections.Generic.HashSet<AST.Address>>;
using InputCell2FormulaDict = System.Collections.Generic.Dictionary<AST.Address, System.Collections.Generic.HashSet<AST.Address>>;

namespace DataDebugMethods
{
    public class DAG
    {
        private Excel.Application _app;
        private CellRefDict _all_cells = new CellRefDict();                 // maps every cell to its COMRef
        private VectorRefDict _all_vectors = new VectorRefDict();           // maps every vector to its COMRef
        private FormulaDict _formulas = new FormulaDict();                  // maps every formula to its formula expr
        private Formula2VectDict _f2v = new Formula2VectDict();             // maps every formula to its input vectors
        private Vect2FormulaDict _v2f = new Vect2FormulaDict();             // maps every input vector to its formulas
        private Formula2InputCellDict _f2i = new Formula2InputCellDict();   // maps every formula to its single-cell inputs
        private Vect2InputCellDict _v2i = new Vect2InputCellDict();         // maps every input vector to its component input cells
        private InputCell2VectDict _i2v = new InputCell2VectDict();         // maps every component input cell to its vectors
        private InputCell2FormulaDict _i2f = new InputCell2FormulaDict();   // maps every single-cell input to its formulas
        private Dictionary<AST.Range, bool> _do_not_perturb = new Dictionary<AST.Range, bool>();    // vector perturbability
        private Dictionary<AST.Address, int> _weights = new Dictionary<AST.Address, int>();         // graph node weight

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
                Excel.Range urng = worksheet.UsedRange;

                // get dimensions
                var left = urng.Column;                      // 1-based left-hand y coordinate
                var right = urng.Columns.Count + left - 1;   // 1-based right-hand y coordinate
                var top = urng.Row;                          // 1-based top x coordinate
                var bottom = urng.Rows.Count + top - 1;      // 1-based bottom x coordinate

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
                object[,] formulas = urng.Formula;

                // for every cell that is actually a formula, add to 
                // formula dictionary & init formula lookup dictionaries
                for (int c = 1; c <= x_max; c++)
                {
                    for (int r = 1; r <= y_max; r++)
                    {
                        var f = (string)formulas[c,r];
                        if (fn_filter.IsMatch(f)) {
                            var addr = AST.Address.NewFromR1C1(r, c, wsname_opt, wbname_opt, path_opt);
                            _formulas.Add(addr, f);
                            _f2v.Add(addr, new HashSet<AST.Range>());
                            _f2i.Add(addr, new HashSet<AST.Address>());
                        }
                    }
                }

                // for each COM object in the used range, create an address object
                // WITHOUT calling any methods on the COM object itself
                foreach (Excel.Range cell in urng)
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

        public AST.COMRef getCOMRefForAddress(AST.Address addr)
        {
            return _all_cells[addr];
        }

        public AST.Address[] getFormulaAddrs()
        {
            return _formulas.Keys.ToArray();
        }

        public AST.COMRef makeInputVectorCOMRef(AST.Range rng)
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
                _do_not_perturb.Add(rng, false);    // initially mark as perturbable
            }
            return c;
        }

        public void linkInputVector(AST.Address formula_addr, AST.Range vector_rng) {
            // add range to range-lookup-by-formula_addr dictionary
            // (initialized in DAG constructor)
            _f2v[formula_addr].Add(vector_rng);
            // add formula_addr to faddr-lookup-by-range dictionary,
            // initializing bucket if necessary
            if (!_v2f.ContainsKey(vector_rng))
            {
                _v2f.Add(vector_rng, new HashSet<AST.Address>());
            }
            if (!_v2f[vector_rng].Contains(formula_addr))
            {
                _v2f[vector_rng].Add(formula_addr);
            }
        }

        public void linkComponentInputCell(AST.Range input_range, AST.Address input_addr)
        {
            // add input_addr to iaddr-lookup-by-input_range dictionary,
            // initializing bucket if necessary
            if (!_v2i.ContainsKey(input_range))
            {
                _v2i.Add(input_range, new HashSet<AST.Address>());
            }
            if (!_v2i[input_range].Contains(input_addr))
            {
                _v2i[input_range].Add(input_addr);
            }
            // add input_range to irng-lookup-by-iaddr dictionary,
            // initializing bucket if necessary
            if (!_i2v.ContainsKey(input_addr))
            {
                _i2v.Add(input_addr, new HashSet<AST.Range>());
            }
            if (!_i2v[input_addr].Contains(input_range))
            {
                _i2v[input_addr].Add(input_range);
            }
        }

        public void linkSingleCellInput(AST.Address formula_addr, AST.Address input_addr)
        {
            // add address to input_addr-lookup-by-formula_addr dictionary
            // (initialzied in DAG constructor)
            _f2i[formula_addr].Add(input_addr);
            // add formula_addr to faddr-lookup-by-iaddr dictionary,
            // initializing bucket if necessary
            if (!_i2f.ContainsKey(input_addr))
            {
                _i2f.Add(input_addr, new HashSet<AST.Address>());
            }
            if (!_i2f[input_addr].Contains(formula_addr))
            {
                _i2f[input_addr].Add(formula_addr);
            }
        }

        public void markPerturbability(AST.Range vector_rng)
        {
            // get inputs
            var inputs = _v2i[vector_rng];

            // count inputs that are formulas
            int fcnt = inputs.Select(iaddr => _formulas.ContainsKey(iaddr)).Count(e => e == true);

            if (fcnt == inputs.Count)
            {
                _do_not_perturb[vector_rng] = true;
            }
        }

        public bool containsLoop()
        {
            throw new NotImplementedException();
        }

        public AST.Address[] terminalFormulaNodes(bool all_outputs)
        {
            // return only the formula nodes that do not serve
            // as input to another cell and that are also not
            // in our list of excluded functions
            if (all_outputs)
            {
                return getFormulaAddrs();
            }
            else
            {
                // start with an array of all formula addresses
                return getFormulaAddrs().Where(addr =>
                    _i2f[addr].Count == 0 &&        // only where the number of formulas consuming this formula == 0
                    _i2v[addr].SelectMany(rng =>    // and where the number of formulas consuming a vector
                        _v2f[rng])                  // containing this formula == 0
                    .Count() == 0     
                ).ToArray();
            }
        }

        public void setWeight(AST.Address node, int weight)
        {
            if (!_weights.ContainsKey(node))
            {
                _weights.Add(node, weight);
            }
            else
            {
                _weights[node] = weight;
            }
        }

        public int getWeight(AST.Address node)
        {
            return _weights[node];
        }

        public HashSet<AST.Range> getFormulaInputVectors(AST.Address f)
        {
            // no need to check for key existence; empty
            // HashSet initialized in DAG constructor
            return _f2v[f];
        }

        internal bool isFormula(AST.Address node)
        {
            return _formulas.ContainsKey(node);
        }

        internal HashSet<AST.Address> getFormulaSingleCellInputs(AST.Address node)
        {
            // no need to check for key existence; empty
            // HashSet initialized in DAG constructor
            return _f2i[node];
        }
    }
}
