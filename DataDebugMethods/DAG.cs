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
        private CellRefDict _all_cells = new CellRefDict();                 // maps every cell (including formulas) to its COMRef
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
        private readonly long _analysis_time;                               // amount of time to run dependence analysis

        public DAG(Excel.Workbook wb, Excel.Application app, bool ignore_parse_errors)
        {
            // start stopwatch
            var sw = new System.Diagnostics.Stopwatch();
            sw.Start();

            // get application reference
            _app = app;

            // bulk read worksheets
            fastFormulaRead(wb);

            // extract references from formulas
            foreach (AST.Address formula_addr in this.getAllFormulaAddrs())
            {
                // get COMRef read earlier
                var formula_ref = this.getCOMRefForAddress(formula_addr);

                foreach (AST.Range vector_rng in ExcelParserUtility.GetRangeReferencesFromFormula(formula_ref, ignore_parse_errors))
                {
                    // fetch/create COMRef, as appropriate
                    var vector_ref = this.makeInputVectorCOMRef(vector_rng);

                    // link formula and input vector
                    this.linkInputVector(formula_addr, vector_rng);

                    // link input vector to the vector's single inputs
                    foreach (AST.Address input_single in vector_rng.Addresses())
                    {
                        this.linkComponentInputCell(vector_rng, input_single);
                    }

                    // if num single inputs = num formulas,
                    // mark vector as non-perturbable
                    this.markPerturbability(vector_rng);
                }

                foreach (AST.Address input_single in ExcelParserUtility.GetSingleCellReferencesFromFormula(formula_ref, ignore_parse_errors))
                {
                    // link formula and single input
                    this.linkSingleCellInput(formula_addr, input_single);
                }
            }

            // stop stopwatch
            sw.Stop();
            _analysis_time = sw.ElapsedMilliseconds;
        }

        private void fastFormulaRead(Excel.Workbook wb)
        {
            // get names once
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
                int width = right - left + 1;
                int height = bottom - top + 1;

                // if the used range is a single cell, Excel changes the type
                if (left == right && top == bottom)
                {
                    var f = (string)urng.Formula;
                    if (fn_filter.IsMatch(f))
                    {
                        var addr = AST.Address.NewFromR1C1(top, left, wsname_opt, wbname_opt, path_opt);
                        _formulas.Add(addr, f);
                        _f2v.Add(addr, new HashSet<AST.Range>());
                        _f2i.Add(addr, new HashSet<AST.Address>());
                    }
                }
                else
                {
                    // array read of formula cells
                    // note that this is a 1-based 2D multiarray
                    object[,] formulas = urng.Formula;

                    // for every cell that is actually a formula, add to 
                    // formula dictionary & init formula lookup dictionaries
                    for (int c = 1; c <= width; c++)
                    {
                        for (int r = 1; r <= height; r++)
                        {
                            var f = (string)formulas[r, c];
                            if (fn_filter.IsMatch(f))
                            {
                                var addr = AST.Address.NewFromR1C1(r + top - 1, c + left - 1, wsname_opt, wbname_opt, path_opt);
                                _formulas.Add(addr, f);
                                _f2v.Add(addr, new HashSet<AST.Range>());
                                _f2i.Add(addr, new HashSet<AST.Address>());
                            }
                        }
                    }
                }

                // for each COM object in the used range, create an address object
                // WITHOUT calling any methods on the COM object itself
                int x = -1;
                int y = 0;
                foreach (Excel.Range cell in urng)
                {
                    // The basic idea here is that we know how Excel iterates over collections
                    // of cells.  The Excel.Range returned by UsedRange is always rectangular.
                    // Thus we can calculate the addresses of each COM cell reference without
                    // needing to incur the overhead of actually asking it for its address.
                    var x_old = x;
                    x = (x + 1) % width;
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

        public string readCOMValueAtAddress(AST.Address addr)
        {
            // null values become the empty string
            var s = System.Convert.ToString(this.getCOMRefForAddress(addr).Range.Value2);
            if (s == null)
            {
                return "";
            }
            else
            {
                return s;
            }
        }

        public long AnalysisMilliseconds
        {
            get { return _analysis_time;  }
        }

        public AST.COMRef getCOMRefForAddress(AST.Address addr)
        {
            return _all_cells[addr];
        }

        public AST.COMRef getCOMRefForRange(AST.Range rng)
        {
            return _all_vectors[rng];
        }

        public string getFormulaAtAddress(AST.Address addr)
        {
            return _formulas[addr];
        }

        public AST.Address[] getAllFormulaAddrs()
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
                _do_not_perturb.Add(rng, true);    // initially mark as not perturbable
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
            int fcnt = inputs.Count(iaddr => _formulas.ContainsKey(iaddr));

            // If there is at least one input that is not a formula
            // mark the whole vector as perturbable.
            // Note: all vectors marked as non-perturbable by default.
            if (fcnt != inputs.Count)
            {
                _do_not_perturb[vector_rng] = false;
            }
        }

        public bool containsLoop()
        {
            var OK = true;
            var visited_from = new Dictionary<AST.Address, AST.Address>();
            foreach (AST.Address addr in _formulas.Keys)
            {
                OK = OK && !traversalHasLoop(addr, visited_from, null);
            }
            return !OK;
        }

        private bool traversalHasLoop(AST.Address current_addr, Dictionary<AST.Address, AST.Address> visited, AST.Address from_addr)
        {
            // base case 1: loop check
            if (visited.ContainsKey(current_addr))
            {
                return true;
            }
            // base case 2: an input cell
            if (!_formulas.ContainsKey(current_addr))
            {
                return false;
            }
            // recursive case (it's a formula)
            // check both single inputs and the inputs of any vector inputs
            bool OK = true;
            HashSet<AST.Address> single_inputs = _f2i[current_addr];
            HashSet<AST.Address> vector_inputs = new HashSet<AST.Address>(_f2v[current_addr].SelectMany(addrs => addrs.Addresses()));
            foreach (AST.Address input_addr in vector_inputs.Union(single_inputs))
            {
                if (OK)
                {
                    // new dict to mark visit
                    var visited2 = new Dictionary<AST.Address, AST.Address>(visited);
                    // mark visit
                    visited2.Add(current_addr, from_addr);
                    // recurse
                    OK = OK && !traversalHasLoop(input_addr, visited2, from_addr);
                }
            }
            return !OK;
        }

        public string ToDOT()
        {
            var visited = new HashSet<AST.Address>();
            String s = "digraph spreadsheet {\n";
            foreach (AST.Address formula_addr in _formulas.Keys)
            {
                s += ToDOT(formula_addr, visited);
            }
            return s + "\n}";
        }

        private string DOTEscapedFormulaString(string formula)
        {
            return formula.Replace("\"", "\\\"");
        }

        private string DOTNodeName(AST.Address addr) {
            return "\"" + addr.A1Local() + "[" + (_formulas.ContainsKey(addr) ? DOTEscapedFormulaString(_formulas[addr]) : readCOMValueAtAddress(addr)) + "]\"";
        }

        private string ToDOT(AST.Address current_addr, HashSet<AST.Address> visited)
        {
            // base case 1: loop protection
            if (visited.Contains(current_addr))
            {
                return "";
            }
            // base case 2: an input
            if (!_formulas.ContainsKey(current_addr))
            {
                return "";
            }
            // case 3: a formula

            String s = "";
            var ca_name = DOTNodeName(current_addr);

            // 3a. single-cell input 
            HashSet<AST.Address> single_inputs = _f2i[current_addr];
            foreach (AST.Address input_addr in single_inputs)
            {
                var ia_name = DOTNodeName(input_addr);

                // print
                s += ia_name + " -> " + ca_name + ";\n";

                // mark visit
                visited.Add(input_addr);

                // recurse
                s += ToDOT(input_addr, visited);
            }

            // 3b. vector input
            HashSet<AST.Range> vector_inputs = _f2v[current_addr];
            foreach (AST.Range v_addr in vector_inputs)
            {
                var rng_name = "\"" + v_addr.GetCOMObject(_app).Address + "\"";

                // print
                s += rng_name + " -> " + ca_name + ";\n";

                // recurse
                foreach (AST.Address input_addr in v_addr.Addresses())
                {
                    var ia_name = DOTNodeName(input_addr);

                    // print
                    s += ia_name + " -> " + rng_name + ";\n";

                    // mark visit
                    visited.Add(input_addr);

                    s += ToDOT(input_addr, visited);
                }
            }

            return s;
        }

        public AST.Address[] terminalFormulaNodes(bool all_outputs)
        {
            // return only the formula nodes that do not serve
            // as input to another cell and that are also not
            // in our list of excluded functions
            if (all_outputs)
            {
                return getAllFormulaAddrs();
            }
            else
            {
                // get all formula addresses
                return getAllFormulaAddrs().Where(addr =>
                    // such that the number of formulas consuming this formula == 0
                    (!_i2f.ContainsKey(addr) || _i2f[addr].Count == 0) &&
                    // and the number of vectors containing this formula == 0
                    (!_i2v.ContainsKey(addr) || _i2v[addr].Count == 0)
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

        public bool isFormula(AST.Address node)
        {
            return _formulas.ContainsKey(node);
        }

        public HashSet<AST.Address> getFormulaSingleCellInputs(AST.Address node)
        {
            // no need to check for key existence; empty
            // HashSet initialized in DAG constructor
            return _f2i[node];
        }

        public AST.Range[] terminalInputVectors()
        {
            return _do_not_perturb.Where(pair => pair.Value == false).Select(pair => pair.Key).ToArray();

            // TODO: it's not clear why the code below returns nothing.

            // this should filter out the following two cases:
            // 1. input range is intermediate (acts as input to a formula
            // and also contains a formula that consumes input from
            // another range).
            // 2. the range is actually a formula cell
            //return _all_vectors.AsTUEnum().Where( pair =>
            //         !pair.Value.DoNotPerturb &&        // range is not marked Do Not Perturb
            //         !pair.Key.Addresses().Any(addr =>  // and range does not contain a formula
            //           _formulas.ContainsKey(addr)
            //         )
            //       ).Select(pair => pair.Key).ToArray();
        }

        public AST.Address[] allComputationCells()
        {
            // get all of the input ranges for all of the functions
            var inputs = _f2v.Values.SelectMany(rngs => rngs.SelectMany(rng => rng.Addresses())).Distinct();

            // get all of the single-cell inputs for all of the functions
            var scinputs = _f2i.Values.SelectMany(rngs => rngs).Distinct();

            // concat all together and return
            return inputs.Concat(scinputs).Distinct().ToArray();
        }

        public AST.Address[] terminalInputCells()
        {
            // this folds all of the inputs for all of the
            // outputs into a set of distinct data-containing cells
            var iecells = terminalFormulaNodes(true).Aggregate(
                              Enumerable.Empty<AST.Address>(),
                              (acc, node) => acc.Union<AST.Address>(getChildCellsRec(node))
                          );
            return iecells.ToArray<AST.Address>();
        }

        private IEnumerable<AST.Address> getChildCellsRec(AST.Address cell_addr)
        {
            // recursive case
            if (_formulas.ContainsKey(cell_addr)) {
                // recursively get vector inputs
                var vector_children = _f2v[cell_addr].SelectMany(rng => getVectorChildCellsRec(rng));

                // recursively get single-cell inputs
                var sc_children = _f2i[cell_addr].SelectMany(cell => getChildCellsRec(cell));

                return vector_children.Concat(sc_children);
            // base case
            } else {
                return new List<AST.Address> { cell_addr };
            }
        }

        private IEnumerable<AST.Address> getVectorChildCellsRec(AST.Range vector_addr)
        {
            // get single-cell inputs (vectors only consist of single cells)
            return _v2i[vector_addr].SelectMany(rng => getChildCellsRec(rng));
        }

        public AST.Range[] allVectors()
        {
            return _all_vectors.KeysT.ToArray();
        }

        public AST.Address[] allCells()
        {
            return _all_cells.KeysT.ToArray();
        }
    }
}
