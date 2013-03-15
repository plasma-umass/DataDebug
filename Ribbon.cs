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
using ColorDict = System.Collections.Generic.Dictionary<Microsoft.Office.Interop.Excel.Workbook, System.Collections.Generic.List<DataDebugMethods.TreeNode>>;

namespace DataDebug
{
    public partial class Ribbon
    {
        private int TRANSPARENT_COLOR_INDEX = -4142;  //-4142 is the transparent default background
        List<TreeNode> originalColorNodes = new List<TreeNode>(); //List for storing the original colors for all nodes
        ColorDict Colors = new ColorDict();  // list for storing colors

        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void DisplayGraphvizTree(AnalysisData analysisData)
        {
            string gvstr = ConstructTree.GenerateGraphVizTree(analysisData.nodes);
            Display disp = new Display();
            disp.textBox1.Text = gvstr;
            disp.ShowDialog();
        }

        // Clear saved colors if the workbook matches
        private void DeleteColorsForWorkbook(ref ColorDict color_storage, Excel.Workbook wb)
        {
            if (color_storage.ContainsKey(wb))
            {
                color_storage.Remove(wb);
            }
        }

        // Save current colors
        private void SaveColors(ref ColorDict color_storage, Excel.Workbook wb)
        {
            List<TreeNode> ts;
            if (!color_storage.TryGetValue(wb, out ts))
            {
                ts = new List<TreeNode>();
                color_storage.Add(wb, ts);
            }

            foreach (Excel.Worksheet ws in wb.Worksheets)
            {
                foreach (Excel.Range cell in ws.UsedRange)
                {
                    //Create a TreeNode for every cell with the name being the cell's address and set the node's worksheet appropriately
                    TreeNode n = new TreeNode(cell.Address, cell.Worksheet, Globals.ThisAddIn.Application.ActiveWorkbook);
                    n.setOriginalColor(System.Drawing.ColorTranslator.FromOle((int)cell.Interior.Color));
                    ts.Add(n);
                }
            }
        }

        // Restore colors to saved value, if we saved them
        private void RestoreColorsForWorkbook(ref ColorDict color_storage, Excel.Workbook wb)
        {
            List<TreeNode> ts;
            if (color_storage.TryGetValue(wb, out ts))
            {
                foreach (TreeNode t in ts)
                {
                    if (!t.isChart() && !t.isRange())
                    {
                        if (!t.getOriginalColor().Equals("Color [White]"))
                        {
                            t.getWorksheetObject().get_Range(t.getName()).Interior.Color = t.getOriginalColor();
                        }
                        else
                        {
                            t.getWorksheetObject().get_Range(t.getName()).Interior.ColorIndex = TRANSPARENT_COLOR_INDEX;
                        }
                    }
                }

                color_storage.Remove(wb);
            }
        }

        //Action for "Analyze Worksheet" button
        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            //Disable screen updating during perturbation and analysis to speed things up
            Globals.ThisAddIn.Application.ScreenUpdating = false;

            // Get current app
            Excel.Application app = Globals.ThisAddIn.Application;

            // Make a new analysisData object
            // TODO If the tool has already run, update the existing instance (so that the colors from the previous run can still be cleared)
            // UPDATE analysisData here
            AnalysisData data = new AnalysisData(Globals.ThisAddIn.Application);
            data.worksheets = app.Worksheets;
            data.global_stopwatch.Reset();
            data.global_stopwatch.Start();

            // Construct a new tree every time the tool is run
            data.Reset();

            // reset colors
            DeleteColorsForWorkbook(ref Colors, app.ActiveWorkbook);

            // save colors
            SaveColors(ref Colors, app.ActiveWorkbook);
            
            // Build dependency graph (modifies data)
            ConstructTree.constructTree(data, app);

            // Perturb data (modifies data)
            Analysis.perturbationAnalysis(data);
            
            // Find outliers (modifies data)
            Analysis.outlierAnalysis(data);

            // Enable screen updating when we're done
            Globals.ThisAddIn.Application.ScreenUpdating = true;
        }

        //Button for testing random code :)
        private void button7_Click(object sender, RibbonControlEventArgs e)
        {
            System.Windows.Forms.MessageBox.Show("Does nothing.");
        }

        //Action for "Clear coloring" button
        private void button8_Click(object sender, RibbonControlEventArgs e)
        {
            RestoreColorsForWorkbook(ref Colors, Globals.ThisAddIn.Application.ActiveWorkbook);
        }

        private void TestNewProcedure_Click(object sender, RibbonControlEventArgs e)
        {
            // Disable screen updating during perturbation and analysis to speed things up
            Globals.ThisAddIn.Application.ScreenUpdating = false;

            // Get current app
            Excel.Application app = Globals.ThisAddIn.Application;

            // Make a new analysisData object
            // TODO If the tool has already run, update the existing instance (so that the colors from the previous run can still be cleared)
            // UPDATE analysisData here
            AnalysisData data = new AnalysisData(Globals.ThisAddIn.Application);
            data.worksheets = app.Worksheets;
            data.global_stopwatch.Reset();
            data.global_stopwatch.Start();

            // Construct a new tree every time the tool is run
            data.Reset();

            // discard any old colors for this workbook
            DeleteColorsForWorkbook(ref Colors, app.ActiveWorkbook);

            // save colors
            SaveColors(ref Colors, app.ActiveWorkbook);

            // Build dependency graph (modifies data)
            ConstructTree.constructTree(data, app);

            // Get bootstraps
            Analysis.Bootstrap(1000, data);

            // Enable screen updating when we're done
            Globals.ThisAddIn.Application.ScreenUpdating = true;
        }
    }
}
