using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
using DataDebugMethods;
using TreeNode = DataDebugMethods.TreeNode;
using CellDict = System.Collections.Generic.Dictionary<AST.Address, string>;
using TreeScore = System.Collections.Generic.Dictionary<DataDebugMethods.TreeNode, int>;
using ErrorDict = System.Collections.Generic.Dictionary<AST.Address, double>;

namespace UserSimulation
{
    public struct PrepData
    {
        public AnalysisData graph;
        public CellDict original_inputs;
        public CellDict correct_outputs;
        public TreeNode[] terminal_input_nodes;
        public TreeNode[] terminal_formula_nodes;
    }

    public static class Prep
    {
        public static PrepData PrepSimulation(Excel.Application app, Excel.Workbook wbh, ProgBar pb)
        {
            // build graph
            var graph = DataDebugMethods.ConstructTree.constructTree(wbh, app);
            if (graph.ContainsLoop())
            {
                throw new DataDebugMethods.ContainsLoopException();
            }
            pb.IncrementProgress(16);

            // get terminal input and terminal formula nodes once
            var terminal_input_nodes = graph.TerminalInputNodes();
            var terminal_formula_nodes = graph.TerminalFormulaNodes(true);  ///the boolean indicates whether to use all outputs or not

            if (terminal_input_nodes.Length == 0)
            {
                throw new NoRangeInputs();
            }

            if (terminal_formula_nodes.Length == 0)
            {
                throw new NoFormulas();
            }

            // save original spreadsheet state
            CellDict original_inputs = UserSimulation.Utility.SaveInputs(graph);

            // force a recalculation before saving outputs, otherwise we may
            // erroneously conclude that the procedure did the wrong thing
            // based solely on Excel floating-point oddities
            UserSimulation.Utility.InjectValues(app, wbh, original_inputs);

            // save function outputs
            CellDict correct_outputs = UserSimulation.Utility.SaveOutputs(terminal_formula_nodes);

            return new PrepData()
            {
                graph = graph,
                original_inputs = original_inputs,
                correct_outputs = correct_outputs,
                terminal_input_nodes = terminal_input_nodes,
                terminal_formula_nodes = terminal_formula_nodes
            };
        }
    }
}
