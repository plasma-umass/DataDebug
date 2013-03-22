using System;
using System.Collections.Generic;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.Office.Tools.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using DataDebugMethods;
using TreeNode = DataDebugMethods.TreeNode;
using ColorDict = System.Collections.Generic.Dictionary<Microsoft.Office.Interop.Excel.Workbook, System.Collections.Generic.List<DataDebugMethods.TreeNode>>;

namespace DataDebug
{
    public partial class Ribbon
    {
        List<TreeNode> originalColorNodes = new List<TreeNode>(); //List for storing the original colors for all nodes
        ColorDict Colors = new ColorDict();  // list for storing colors

        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        // Action for "Analyze Worksheet" button
        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            //Disable screen updating during perturbation and analysis to speed things up
            Globals.ThisAddIn.Application.ScreenUpdating = false;

            // Get current app
            Excel.Application app = Globals.ThisAddIn.Application;

            // Make a new analysisData object
            AnalysisData data = new AnalysisData(Globals.ThisAddIn.Application);
            data.worksheets = app.Worksheets;
            data.global_stopwatch.Reset();
            data.global_stopwatch.Start();

            // Construct a new tree every time the tool is run
            data.Reset();

            // reset colors
            RibbonHelper.DeleteColorsForWorkbook(ref Colors, app.ActiveWorkbook);

            // save colors
            RibbonHelper.SaveColors(ref Colors, app.ActiveWorkbook);
            
            // Build dependency graph (modifies data)
            ConstructTree.constructTree(data, app);

            // Perturb data (modifies data)
            Analysis.perturbationAnalysis(data);
            
            // Find outliers (modifies data)
            Analysis.outlierAnalysis(data);

            // Enable screen updating when we're done
            Globals.ThisAddIn.Application.ScreenUpdating = true;
        }

        // Button for testing random code :)
        private void button7_Click(object sender, RibbonControlEventArgs e)
        {
            System.Windows.Forms.MessageBox.Show("Does nothing.");
        }

        // Action for "Clear coloring" button
        private void button8_Click(object sender, RibbonControlEventArgs e)
        {
            RibbonHelper.RestoreColorsForWorkbook(ref Colors, Globals.ThisAddIn.Application.ActiveWorkbook);
        }

        private void TestNewProcedure_Click(object sender, RibbonControlEventArgs e)
        {
            // Disable screen updating during perturbation and analysis to speed things up
            Globals.ThisAddIn.Application.ScreenUpdating = false;

            // Get current app
            Excel.Application app = Globals.ThisAddIn.Application;

            // Make a new analysisData object
            AnalysisData data = new AnalysisData(Globals.ThisAddIn.Application);
            data.worksheets = app.Worksheets;
            data.global_stopwatch.Reset();
            data.global_stopwatch.Start();

            // Construct a new tree every time the tool is run
            data.Reset();

            // discard any old colors for this workbook
            RibbonHelper.DeleteColorsForWorkbook(ref Colors, app.ActiveWorkbook);

            // save colors
            RibbonHelper.SaveColors(ref Colors, app.ActiveWorkbook);

            // Build dependency graph (modifies data)
            ConstructTree.constructTree(data, app);

            // Get bootstraps
            Analysis.Bootstrap(1000, data);

            // Enable screen updating when we're done
            Globals.ThisAddIn.Application.ScreenUpdating = true;
        }
    }
}
