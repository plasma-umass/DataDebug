using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.Office.Tools.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using System.Text.RegularExpressions;
using DataDebugMethods;
using Microsoft.FSharp.Core;
using TreeNode = DataDebugMethods.TreeNode;
using System.Diagnostics;
using TreeDict = System.Collections.Generic.Dictionary<AST.Address, DataDebugMethods.TreeNode>;
using TreeDictPair = System.Collections.Generic.KeyValuePair<AST.Address, DataDebugMethods.TreeNode>;

namespace DataDebug
{
    public partial class Ribbon
    {
        private int TRANSPARENT_COLOR_INDEX = -4142;  //-4142 is the transparent default background
        //private bool toolHasNotRun = true; //this is to keep track of whether the tool has already run without having cleared the colorings
        List<TreeNode> originalColorNodes = new List<TreeNode>(); //List for storing the original colors for all nodes
        //List<TreeNode> nodelist;        //This is a list holding all the TreeNodes in the Excel file
        //double[][][][] impacts_grid; //This is a multi-dimensional array of doubles that will hold each cell's impact on each of the outputs
        //bool[][][][] reachable_grid; //This is a multi-dimensional array of bools that will indicate whether a certain output is reachable from a certain cell
        //double[][] min_max_delta_outputs; //This keeps the min and max delta for each output; first index indicates the output index; second index 0 is the min delta, 1 is the max delta for that output
        //List<TreeNode> ranges;      // This is a list of input ranges, with each Excel.Range COM object encapsulated in a TreeNode
        //List<StartValue> starting_outputs; //This will store the values of all the output nodes at the start of the procedure for swapping values (fuzzing)
        //List<TreeNode> output_cells; //This will store all the output nodes at the start of the fuzzing procedure
        //List<double[]>[] reachable_impacts_grid;  //This will store impacts for cells reachable from a particular output
        //double[][][] reachable_impacts_grid_array; //This will store impacts for cells reachable from a particular output in array form
        //int input_cells_in_computation_count = 0;
        //int raw_input_cells_in_computation_count = 0;
        //int formula_cells_count;
        //System.Diagnostics.Stopwatch global_stopwatch = new System.Diagnostics.Stopwatch();
        //ProgBar pb;
        //TreeDict nodes;
        //TimeSpan tree_building_timespan;
        //TimeSpan impact_scoring_timespan;
        //TimeSpan swapping_timespan;
        //int outliers_count; //This gets assigned and updated in the Analysis class
        //int[][][] times_perturbed;

        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        // TODO move analysis into separate lib

        /*
         * This method constructs the dependency graph from the worksheet.
         * It analyzes formulas and looks for references to cells or ranges of cells.
         * It also looks for any charts, and adds those to the dependency graph as well. 
         * After the dependency graph is constructed, we use it to determine and propagate weights to all nodes in the graph. 
         * This method also contains the perturbation procedure and outlier analysis logic.
         * In the end, a text representation of the dependency graph is given in GraphViz format. It includes the entire graph and the weights of the nodes.
         */
        private void constructTree(AnalysisData analysisData)
        {
            analysisData.pb.SetProgress(0);
            analysisData.impact_scoring_timespan = analysisData.global_stopwatch.Elapsed;
            analysisData.swapping_timespan = analysisData.global_stopwatch.Elapsed;
            analysisData.input_cells_in_computation_count = 0;
            analysisData.raw_input_cells_in_computation_count = 0;

            // Get a range representing the formula cells for each worksheet in each workbook
            ArrayList formulaRanges = ConstructTree.GetFormulaRanges(Globals.ThisAddIn.Application.Worksheets, Globals.ThisAddIn.Application);
            analysisData.formula_cells_count = ConstructTree.CountFormulaCells(formulaRanges);
            
            // Create nodes for every cell containing a formula
            analysisData.nodes = ConstructTree.CreateFormulaNodes(formulaRanges, Globals.ThisAddIn.Application);
            
            //Now we parse the formulas in nodes to extract any range and cell references
            int node_count = analysisData.nodes.Count; // we save this because nodes.Count grows in this loop
            for (int nodeIndex = 0; nodeIndex < node_count; nodeIndex++)
            {
                TreeNode node = analysisData.nodes.ElementAt(nodeIndex).Value; // nodePair.Value;

                // For each of the ranges found in the formula by the parser, do the following:
                foreach (Excel.Range range in ExcelParserUtility.GetReferencesFromFormula(node.getFormula(), node.getWorkbookObject(), node.getWorksheetObject()))
                {
                    // Make a TreeNode for the Excel range COM object
                    TreeNode rangeNode = ConstructTree.MakeRangeTreeNode(analysisData.ranges, range, node);
                    // Create TreeNodes for each range's Cell and add them as
                    // parents to THIS range's TreeNode
                    ConstructTree.CreateCellNodesFromRange(rangeNode, node, analysisData.nodes);
                }
            }

            //TODO -- we are not able to capture ranges that are identified in stored procedures or macros, just ones referenced in formulas
            //TODO -- Dealing with fuzzing of charts -- idea: any cell that feeds into a chart is essentially an output; the chart is just a visual representation (can charts operate on values before they are displayed? don't think so...)
            analysisData.starting_outputs = new List<StartValue>(); //This will store the values of all the output nodes at the start of the procedure for swapping values (fuzzing)
            analysisData.output_cells = new List<TreeNode>(); //This will store all the output nodes at the start of the fuzzing procedure

            ConstructTree.StoreOutputs(analysisData);
            
            //Tree building stopwatch
            analysisData.tree_building_timespan = analysisData.global_stopwatch.Elapsed;
        }

        private void DisplayGraphvizTree(AnalysisData analysisData)
        {
            string gvstr = ConstructTree.GenerateGraphVizTree(analysisData.nodes);
            Display disp = new Display();
            disp.textBox1.Text = gvstr;
            disp.ShowDialog();
        }

            /*
            //Procedure for swapping values within ranges, replacing all repeated values at once
            if (checkBox2.Checked) //Checks if the option for swapping values simultaneously is checked
            {
                List<TreeNode> swap_domain;
                swap_domain = ranges;
                
                foreach (TreeNode node in swap_domain)
                {
                    bool all_children_are_charts = true;
                    if (node.isRange() && node.hasChildren())
                    {
                        //bool children_are_charts = false;
                        foreach (TreeNode child in node.getChildren())
                        {
                            if (!child.isChart())
                            {
                                all_children_are_charts = false;
                            }
                        }
                    }
                    //For each range node, do the following:
                    if (node.isRange() && !all_children_are_charts)
                    {
                        double[] influences = new double[node.getParents().Count];  //Array to keep track of the influence values for every cell
                        int influence_index = 0;        //Keeps track of the current position in the influences array
                        double max_total_delta = 0;     //The maximum influence found (for normalizing)
                        double min_total_delta =-1;     //The minimum influence found (for normalizing)
                        //Swapping values; loop over all nodes in the range
                        foreach (TreeNode parent in node.getParents())
                        {
                            String twin_cells_string = parent.getName();
                            //Find any nodes with a matching value and keep track of them
                            int twin_count = 1;     //This will keep track of the number of cells that have this exact value
                            foreach (TreeNode twin in node.getParents())
                            {
                                if (twin.getName() == parent.getName()) // if twin is the same cell as the current cell being examined, don't do anything
                                {
                                    continue;
                                }
                                if (twin.getWorksheetObject().get_Range(twin.getName()).Value == parent.getWorksheetObject().get_Range(parent.getName()).Value)
                                {
                                    twin_cells_string = twin_cells_string + "," + twin.getName();
                                    twin_count++;
                                }
                            }
                            Excel.Range twin_cells = parent.getWorksheetObject().get_Range(twin_cells_string);
                            String[] formulas = new String[twin_count]; //Stores the formulas in the twin_cells
                            int i = 0; //Counter for indexing within the formulas array
                            foreach (Excel.Range cell in twin_cells)
                            {
                                if (cell.HasFormula)
                                {
                                    formulas[i] = cell.Formula;
                                }
                                i++;
                            }
                            double start_value = parent.getWorksheetObject().get_Range(parent.getName()).Value;
                            double total_delta = 0;
                            double delta = 0;
                            foreach (TreeNode sibling in node.getParents())
                            {
                                if (sibling.getName() == parent.getName())
                                {
                                    continue;
                                }
                                Excel.Range sibling_cell = sibling.getWorksheetObject().get_Range(sibling.getName());
                                twin_cells.Value = sibling_cell.Value;
                                int index = 0;
                                delta = 0;
                                foreach (TreeNode n in output_cells)
                                {
                                    if (starting_outputs[index].get_double() != 0)
                                    {
                                        delta = Math.Abs(starting_outputs[index].get_double() - n.getWorksheetObject().get_Range(n.getName()).Value) / Math.Abs(starting_outputs[index].get_double());
                                    }
                                    else
                                    {
                                        delta = Math.Abs(starting_outputs[index].get_double() - n.getWorksheetObject().get_Range(n.getName()).Value);
                                    }
                                    index++;
                                    total_delta = total_delta + delta;
                                }
                            }
                            total_delta = total_delta / (node.getParents().Count - 1);
                            total_delta = total_delta / twin_count;
                            influences[influence_index] = total_delta;
                            influence_index++;
                            if (max_total_delta < total_delta)
                            {
                                max_total_delta = total_delta;
                            }
                            if (min_total_delta > total_delta || min_total_delta == -1)
                            {
                                min_total_delta = total_delta;
                            }
                            twin_cells.Value = start_value;
                            int j = 0;
                            foreach (Excel.Range cell in twin_cells)
                            {
                                if (formulas[j] != null)
                                    cell.Formula = formulas[j];
                                j++;
                            }
                            twin_cells.Interior.Color = System.Drawing.Color.Beige;
                        }
                        int ind = 0;
                        foreach (TreeNode parent in node.getParents())
                        {
                            if (max_total_delta != 0)
                            {
                                influences[ind] = (influences[ind] - min_total_delta) / max_total_delta;
                            }
                            ind++;
                        }
                        int indexer = 0;
                        foreach (TreeNode parent in node.getParents())
                        {
                            Excel.Range cell = parent.getWorksheetObject().get_Range(parent.getName());
                            cell.Interior.Color = System.Drawing.Color.FromArgb(Convert.ToInt32(255 - influences[indexer] * 255), 255, 255);
                            indexer++;
                        }
                    }
                }
            }
            */
        
        //Action for "Analyze Worksheet" button
        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            //If tool is running for the first time, make a new analysisData object
            AnalysisData analysisData = new AnalysisData(Globals.ThisAddIn.Application);
            //If the tool has already run, update the existing instance (so that the colors from the previous run can still be cleared)
                //UPDATE analysisData here
            analysisData.worksheets = Globals.ThisAddIn.Application.Worksheets;
            analysisData.global_stopwatch.Reset();
            analysisData.global_stopwatch.Start();

            //Construct a new tree every time the tool is run
            analysisData.nodelist = new List<TreeNode>();        //This is a list holding all the TreeNodes in the Excel file
            analysisData.ranges = new List<TreeNode>();        //This is a list holding all the ranges of TreeNodes in the Excel file
            
            for (int i = 0; i < originalColorNodes.Count; i++)
            {
                if (originalColorNodes[i].getWorkbookObject() == Globals.ThisAddIn.Application.ActiveWorkbook)
                {
                    originalColorNodes.RemoveAt(i);
                }
            }
            foreach (Excel.Worksheet worksheet in Globals.ThisAddIn.Application.Worksheets)
            {
                foreach (Excel.Range cell in worksheet.UsedRange)
                {
                    TreeNode n = new TreeNode(cell.Address, cell.Worksheet, Globals.ThisAddIn.Application.ActiveWorkbook);  //Create a TreeNode for every cell with the name being the cell's address and set the node's worksheet appropriately
                    //n.setOriginalColor(cell.Interior.ColorIndex);
                    n.setOriginalColor(System.Drawing.ColorTranslator.FromOle((int)cell.Interior.Color));
                    originalColorNodes.Add(n);
                }
            }
            analysisData.pb = new ProgBar(0, 100);
            
            constructTree(analysisData);

            //Disable screen updating during perturbation and analysis to speed things up
            Globals.ThisAddIn.Application.ScreenUpdating = false;

            Analysis.perturbationAnalysis(analysisData);
            
            Analysis.outlierAnalysis(analysisData);
            //Enable screen updating when we're done
            Globals.ThisAddIn.Application.ScreenUpdating = true;
        }

        //Button for testing random code :)
        private void button7_Click(object sender, RibbonControlEventArgs e)
        {
            //System.Windows.Forms.MessageBox.Show(Globals.ThisAddIn.Application.ActiveWorkbook.Path + "");
            //System.Windows.Forms.MessageBox.Show(Globals.ThisAddIn.Application.Workbooks[1] + "");
            //foreach (Excel.Chart chart in Globals.ThisAddIn.Application.Charts)
            //{
            //    foreach (Excel.Series series in (Excel.SeriesCollection)chart.SeriesCollection(Type.Missing))
            //    {
            //        string formula = series.Formula;  //The formula contained in the cell
            //        System.Windows.Forms.MessageBox.Show(formula);
            //    }
            //}
            //ProgBar pb = new ProgBar(0, 100);
            //pb.Show();
        }

        //Action for "Clear coloring" button
        private void button8_Click(object sender, RibbonControlEventArgs e)
        {
            for (int i = 0; i < originalColorNodes.Count; i++)
            {
                if (originalColorNodes[i].getWorkbookObject() == Globals.ThisAddIn.Application.ActiveWorkbook)
                {
                    //If the node is just a cell, clear any coloring
                    if (!originalColorNodes[i].isChart() && !originalColorNodes[i].isRange())
                    {
                        //System.Windows.Forms.MessageBox.Show(node.getName() + " " + node.getOriginalColor());
                        //node.getWorksheetObject().get_Range(node.getName()).Interior.ColorIndex = 0;
                        //node.getWorksheetObject().get_Range(node.getName()).Interior.ColorIndex = node.getOriginalColor();
                        if (!(originalColorNodes[i].getOriginalColor() + "").Equals("Color [White]"))
                        {
                            originalColorNodes[i].getWorksheetObject().get_Range(originalColorNodes[i].getName()).Interior.Color = originalColorNodes[i].getOriginalColor();
                        }
                        else
                        {
                            originalColorNodes[i].getWorksheetObject().get_Range(originalColorNodes[i].getName()).Interior.ColorIndex = TRANSPARENT_COLOR_INDEX;  //-4142 is the transparent default background for cells
                        }
                    }
                    originalColorNodes.RemoveAt(i);
                    i--;
                }
            }
        }

        private void TestParser_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;
            Excel.Worksheet ws = Globals.ThisAddIn.Application.ActiveSheet;
            Excel.Range fn_cell = Globals.ThisAddIn.Application.ActiveCell;
            if (fn_cell.HasFormula)
            {
                string formula = Convert.ToString(fn_cell.Formula);
                IEnumerable<Excel.Range> ranges = ExcelParserUtility.GetReferencesFromFormula(formula, wb, ws);
                foreach (Excel.Range range in ranges)
                {
                    foreach (Excel.Range cell in range)
                    {
                        cell.Interior.Color = System.Drawing.Color.Red;
                    }
                }
            }
            else
            {
                System.Windows.Forms.MessageBox.Show("The currently selected cell is not a formula.");
            }
            
        }
    }
}
