using System;
using System.Collections.Generic;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.Office.Tools.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using DataDebugMethods;
using TreeNode = DataDebugMethods.TreeNode;
using TreeScore = System.Collections.Generic.Dictionary<DataDebugMethods.TreeNode, int>;
using ColorDict = System.Collections.Generic.Dictionary<Microsoft.Office.Interop.Excel.Workbook, System.Collections.Generic.List<DataDebugMethods.TreeNode>>;
using Microsoft.FSharp.Core;
using System.IO;

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

        // Button for outputting MTurk HIT CSVs
        private void button7_Click(object sender, RibbonControlEventArgs e)
        {
            // the longest input field we can represent on MTurk
            const int MAXLEN = 20;

            // get MTurk jobs or fail is spreadsheet data cells are too long
            TurkJob[] turkjobs;
            var turkjobs_opt = ConstructTree.DataForMTurk(Globals.ThisAddIn.Application, MAXLEN);
            if (FSharpOption<TurkJob[]>.get_IsSome(turkjobs_opt))
            {
                turkjobs = turkjobs_opt.Value;
            }
            else
            {
                System.Windows.Forms.MessageBox.Show("This spreadsheet contains data items with lengths longer than " + MAXLEN + " characters and cannot be converted into an MTurk job.");
                return;
            }

            // get workbook name
            var wbname = Globals.ThisAddIn.Application.ActiveWorkbook.Name;

            // prompt for folder name
            var sFD = new System.Windows.Forms.FolderBrowserDialog();
            sFD.ShowDialog();

            // If the path is not an empty string, go ahead
            if (sFD.SelectedPath != "")
            {
                // write file
                var outfile = Path.Combine(sFD.SelectedPath, wbname + ".arr");
                TurkJob.SerializeArray(outfile, turkjobs);

                // write images, 2 for each TurkJob
                foreach (TurkJob tj in turkjobs)
                {
                    tj.WriteAsImages(sFD.SelectedPath, wbname);
                }

                //// sanity check
                //TurkJob[] fromfile = TurkJob.DeserializeArray(saveFileDialog1.FileName);
                //string csv = "job_id,cell1,cell2,cell3,cell4,cell5,cell6,cell7,cell8,cell9,cell10\n";
                //foreach (TurkJob job in turkjobs)
                //{
                //    csv += job.ToCSVLine();
                //}
                //System.Windows.Forms.MessageBox.Show("This is what I got back:\n\n" + csv);
            }
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
            var scores = Analysis.Bootstrap((int)(Math.Ceiling(1000 * Math.Exp(1.0))), data, this.weighted.Checked);

            // Color outputs
            Analysis.ColorOutputs(scores);

            // Enable screen updating when we're done
            Globals.ThisAddIn.Application.ScreenUpdating = true;
        }
    }
}
