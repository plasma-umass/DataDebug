using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;
using DataDebugMethods;

namespace DataDebug
{
    public partial class PerformanceExperiments : Form
    {
        string folderPath; 
        public PerformanceExperiments()
        {
            InitializeComponent();
        }

        private void selectFolder_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog selectFolderDialog = new FolderBrowserDialog();
            selectFolderDialog.ShowDialog();
            if (selectFolderDialog.SelectedPath == "")
            {
                return;
            }
            //a folder was chosen 
            folderPath = @selectFolderDialog.SelectedPath;
            //textBox1.Text += Environment.NewLine + "Folder was selected: " + folderPath;
            textBox1.AppendText(Environment.NewLine + "Folder was selected: " + folderPath);
            //textBox1.Text += Environment.NewLine + "Checking for necessary files";
            textBox1.AppendText(Environment.NewLine + "Checking for files");
            //Look for xls or xlsx
            string[] xlsFilePaths = Directory.GetFiles(folderPath, "*.xls");
            //sstring[] xlsxFilePaths = Directory.GetFiles(folderPath, "*.xlsx");
            if (xlsFilePaths.Length == 0)
            {
                //textBox1.Text += Environment.NewLine + "ERROR: XLS/XLSX file not found";
                textBox1.AppendText(Environment.NewLine + "ERROR: No *.xls or *.xlsx files found.");
                return;
            }
        } //end selectFolder_Click

        private void runExperiments_Click(object sender, EventArgs e)
        {
            string[] xlsFilePaths = Directory.GetFiles(folderPath, "*.xls");
            string results = "Workbook name" + "\tBootstraps" + "\tTotal Time" + "\tTree Building Time" + "\tBootstrap Time" +
                    "\tColoring Time" + Environment.NewLine;
            foreach (string xlsFilePath in xlsFilePaths)
            {
                System.Diagnostics.Stopwatch stopwatch = new System.Diagnostics.Stopwatch();
                TimeSpan tree_building_timespan;
                TimeSpan bootstrap_timespan;
                TimeSpan coloring_timespan;
                TimeSpan total_timespan;

                textBox1.AppendText(Environment.NewLine + "Opening Excel file: " + xlsFilePath + Environment.NewLine);
                
                // Get current app
                Excel.Application app = Globals.ThisAddIn.Application;
                Excel.Workbook originalWB = app.Workbooks.Open(xlsFilePath, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);
                
                textBox1.AppendText(Environment.NewLine + "Running bootstrap analysis." + Environment.NewLine);
                
                //Disable screen updating during perturbation and analysis to speed things up
                Globals.ThisAddIn.Application.ScreenUpdating = false;

                // Make a new analysisData object
                AnalysisData data = new AnalysisData(Globals.ThisAddIn.Application);
                data.worksheets = app.Worksheets;
                
                // Construct a new tree every time the tool is run
                data.Reset();
                stopwatch.Start();
                
                // Build dependency graph (modifies data)
                ConstructTree.constructTree(data, app);
                
                tree_building_timespan = stopwatch.Elapsed;
                string tree_building_time = tree_building_timespan.TotalSeconds + "";

                if (data.TerminalInputNodes().Length == 0)
                {
                    System.Windows.Forms.MessageBox.Show("This spreadsheet has no input ranges.  Sorry, dude.");
                    data.pb.Close();
                    Globals.ThisAddIn.Application.ScreenUpdating = true;
                    return;
                }

                // e * bootstrapMultiplier
                int bootstrapMultiplier = (int)numericUpDown1.Value;
                var NBOOTS = (int)(Math.Ceiling(bootstrapMultiplier * Math.Exp(1.0)));

                // Get bootstraps
                var scores = Analysis.Bootstrap(NBOOTS, data, app, true);

                bootstrap_timespan = stopwatch.Elapsed;
                string bootstrap_time = (bootstrap_timespan.TotalSeconds - tree_building_timespan.TotalSeconds) + ""; 

                // Color outputs
                Analysis.ColorOutputs(scores);
                
                stopwatch.Stop();
                total_timespan = stopwatch.Elapsed;
                string total_time = total_timespan.TotalSeconds + "";
                coloring_timespan = stopwatch.Elapsed;
                string coloring_time = (coloring_timespan.TotalSeconds - bootstrap_timespan.TotalSeconds) + ""; 

                // Enable screen updating when we're done
                Globals.ThisAddIn.Application.ScreenUpdating = true;
                
                textBox1.AppendText("Done." + Environment.NewLine);
                textBox1.AppendText("Total time:" + total_time + 
                    Environment.NewLine + "Tree construction: " + tree_building_time + 
                    Environment.NewLine + "Perturbation: " + bootstrap_time + 
                    Environment.NewLine + "Coloring: " + coloring_time + Environment.NewLine + Environment.NewLine);
                
                results += originalWB.Name + "\t" + NBOOTS + "\t" + total_time + "\t" + tree_building_time + "\t" + bootstrap_time +
                    "\t" + coloring_time + Environment.NewLine;

                originalWB.Close(false);
            }
            System.IO.File.WriteAllText(@folderPath + @"\ExperimentalResults.xls", results);
        }   //END runExperiments_Click

        private void runSingle_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.ShowDialog();
            if (openFileDialog.FileName == "")
            {
                System.Windows.Forms.MessageBox.Show("No file selected.");
            }
            else if (!openFileDialog.FileName.Substring(openFileDialog.FileName.LastIndexOf(".")).Contains(".xls"))
            {
                System.Windows.Forms.MessageBox.Show("Please select .xls or .xlsx file.");
            }
            else
            {
                System.Diagnostics.Stopwatch stopwatch = new System.Diagnostics.Stopwatch();
                TimeSpan tree_building_timespan;
                TimeSpan bootstrap_timespan;
                TimeSpan coloring_timespan;
                TimeSpan total_timespan;

                textBox1.AppendText(Environment.NewLine + "Opening Excel file: " + openFileDialog.FileName + Environment.NewLine);

                // Get current app
                Excel.Application app = Globals.ThisAddIn.Application;
                Excel.Workbook originalWB = app.Workbooks.Open(openFileDialog.FileName, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);

                textBox1.AppendText(Environment.NewLine + "Running bootstrap analysis." + Environment.NewLine);

                //Disable screen updating during perturbation and analysis to speed things up
                Globals.ThisAddIn.Application.ScreenUpdating = false;

                // Make a new analysisData object
                AnalysisData data = new AnalysisData(Globals.ThisAddIn.Application);
                data.worksheets = app.Worksheets;

                // Construct a new tree every time the tool is run
                data.Reset();
                stopwatch.Start();

                // Build dependency graph (modifies data)
                ConstructTree.constructTree(data, app);

                tree_building_timespan = stopwatch.Elapsed;
                string tree_building_time = tree_building_timespan.TotalSeconds + "";

                if (data.TerminalInputNodes().Length == 0)
                {
                    System.Windows.Forms.MessageBox.Show("This spreadsheet has no input ranges.  Sorry, dude.");
                    data.pb.Close();
                    Globals.ThisAddIn.Application.ScreenUpdating = true;
                    return;
                }

                // e * bootstrapMultiplier
                int bootstrapMultiplier = (int)numericUpDown1.Value;
                var NBOOTS = (int)(Math.Ceiling(bootstrapMultiplier * Math.Exp(1.0)));

                // Get bootstraps
                var scores = Analysis.Bootstrap(NBOOTS, data, app, true);

                bootstrap_timespan = stopwatch.Elapsed;
                string bootstrap_time = (bootstrap_timespan.TotalSeconds - tree_building_timespan.TotalSeconds) + "";

                // Color outputs
                Analysis.ColorOutputs(scores);

                stopwatch.Stop();
                total_timespan = stopwatch.Elapsed;
                string total_time = total_timespan.TotalSeconds + "";
                coloring_timespan = stopwatch.Elapsed;
                string coloring_time = (coloring_timespan.TotalSeconds - bootstrap_timespan.TotalSeconds) + "";

                // Enable screen updating when we're done
                Globals.ThisAddIn.Application.ScreenUpdating = true;

                textBox1.AppendText("Done." + Environment.NewLine);
                textBox1.AppendText("Total time:" + total_time +
                    Environment.NewLine + "Tree construction: " + tree_building_time +
                    Environment.NewLine + "Perturbation: " + bootstrap_time +
                    Environment.NewLine + "Coloring: " + coloring_time + Environment.NewLine + Environment.NewLine);

                string results = originalWB.Name + "\t" + NBOOTS + "\t" + total_time + "\t" + tree_building_time + "\t" + bootstrap_time +
                    "\t" + coloring_time + Environment.NewLine;

                originalWB.Close(false);
                string originalFileText = System.IO.File.ReadAllText(@folderPath + @"\ExperimentalResults.xls");
                if (System.IO.File.ReadAllLines(@folderPath + @"\ExperimentalResults.xls").Length > 1)
                {
                    results = originalFileText + results;
                }
                else
                {
                    results = "Workbook name" + "\tBootstraps" + "\tTotal Time" + "\tTree Building Time" + "\tBootstrap Time" +
                    "\tColoring Time" + Environment.NewLine + results;
                }
                System.IO.File.WriteAllText(@folderPath + @"\ExperimentalResults.xls", results);
            }
        }

    }
}
