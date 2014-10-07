using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
using DataDebugMethods;
using CellDict = System.Collections.Generic.Dictionary<AST.Address, string>;
using TreeScore = System.Collections.Generic.Dictionary<AST.Address, int>;
using ErrorDict = System.Collections.Generic.Dictionary<AST.Address, double>;

namespace UserSimulation
{
    public struct PrepData
    {
        public DAG dag;
        public CellDict original_inputs;
        public CellDict correct_outputs;
        public AST.Range[] terminal_input_nodes;
        public AST.Address[] terminal_formula_nodes;
    }

    public static class Prep
    {
        public static PrepData PrepSimulation(Excel.Application app, Excel.Workbook wbh, ProgBar pb, bool ignore_parse_errors)
        {
            // build graph
            var dag = DataDebugMethods.DependenceAnalysis.constructDAG(wbh, app, ignore_parse_errors);
            if (dag.containsLoop())
            {
                throw new DataDebugMethods.ContainsLoopException();
            }
            pb.IncrementProgress(16);

            // get terminal input and terminal formula nodes once
            var terminal_input_nodes = dag.terminalInputVectors();
            var terminal_formula_nodes = dag.terminalFormulaNodes(true);  ///the boolean indicates whether to use all outputs or not

            if (terminal_input_nodes.Length == 0)
            {
                throw new NoRangeInputs();
            }

            if (terminal_formula_nodes.Length == 0)
            {
                throw new NoFormulas();
            }

            // save original spreadsheet state
            CellDict original_inputs = UserSimulation.Utility.SaveInputs(dag);

            // force a recalculation before saving outputs, otherwise we may
            // erroneously conclude that the procedure did the wrong thing
            // based solely on Excel floating-point oddities
            UserSimulation.Utility.InjectValues(app, wbh, original_inputs);

            // save function outputs
            CellDict correct_outputs = UserSimulation.Utility.SaveOutputs(terminal_formula_nodes, dag);

            return new PrepData()
            {
                dag = dag,
                original_inputs = original_inputs,
                correct_outputs = correct_outputs,
                terminal_input_nodes = terminal_input_nodes,
                terminal_formula_nodes = terminal_formula_nodes
            };
        }
    }
}
