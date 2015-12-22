using System;
using System.Collections.Generic;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.Office.Tools.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using Depends;
using DataDebugMethods;
using TreeScore = System.Collections.Generic.Dictionary<AST.Address, int>;
using TreeDict = System.Collections.Generic.Dictionary<AST.Address, AST.Address>;
using Microsoft.FSharp.Core;
using System.IO;
using System.Linq;
using OptTuple = Microsoft.FSharp.Core.FSharpOption<System.Tuple<UserSimulation.Classification, string>>;

namespace CheckCell
{
    public partial class Ribbon
    {
        Dictionary<Excel.Workbook, WorkbookState> wbstates = new Dictionary<Excel.Workbook, WorkbookState>();
        WorkbookState current_workbook;

        // simulation files
        string classification_file;
        String benchmark_dir;
        String simulation_output_dir;
        String simulation_classification_file;

        private void SetUIState(WorkbookState wbs)
        {
            this.MarkAsOKButton.Enabled = wbs.MarkAsOK_Enabled;
            this.FixErrorButton.Enabled = wbs.FixError_Enabled;
            this.StartOverButton.Enabled = wbs.ClearColoringButton_Enabled;
            this.AnalyzeButton.Enabled = wbs.Analyze_Enabled;
        }

        private void SetUIStateNoWorkbooks()
        {
            this.MarkAsOKButton.Enabled = false;
            this.FixErrorButton.Enabled = false;
            this.StartOverButton.Enabled = false;
            this.AnalyzeButton.Enabled = false;
        }

        private void Ribbon_Load(object sender, RibbonUIEventArgs e)
        {
            // Callbacks for handling workbook state objects
            //WorkbookOpen(Globals.ThisAddIn.Application.ActiveWorkbook);
            //((Excel.AppEvents_Event)Globals.ThisAddIn.Application).NewWorkbook += WorkbookOpen;
            Globals.ThisAddIn.Application.WorkbookOpen += WorkbookOpen;
            Globals.ThisAddIn.Application.WorkbookActivate += WorkbookActivated;
            Globals.ThisAddIn.Application.WorkbookDeactivate += WorkbookDeactivated;
            Globals.ThisAddIn.Application.WorkbookBeforeClose += WorkbookClose;

            // sometimes the default blank workbook opens *before* the CheckCell
            // add-in is loaded so we have to handle sheet state specially.
            if (current_workbook == null)
            {
                var wb = Globals.ThisAddIn.Application.ActiveWorkbook;
                if (wb == null)
                {
                    // the plugin loaded first; there's no active workbook
                    return;
                }
                WorkbookOpen(wb);
                WorkbookActivated(wb);
            }
        }

        // This event is called when Excel opens a workbook
        private void WorkbookOpen(Excel.Workbook workbook)
        {
            wbstates.Add(workbook, new WorkbookState(Globals.ThisAddIn.Application, workbook));
        }

        // This event is called when Excel brings an opened workbook
        // to the foreground
        private void WorkbookActivated(Excel.Workbook workbook)
        {
            // when opening a blank sheet, Excel does not emit
            // a WorkbookOpen event, so we need to call it manually
            if (!wbstates.ContainsKey(workbook)) {
                WorkbookOpen(workbook);
            }
            current_workbook = wbstates[workbook];
            SetUIState(current_workbook);
        }

        // This even it called when Excel sends an opened workbook
        // to the background
        private void WorkbookDeactivated(Excel.Workbook workbook)
        {
            current_workbook = null;
            // WorkbookBeforeClose event does not fire for default workbooks
            // containing no data
            var wbs = new List<Excel.Workbook>();
            foreach (var wb in Globals.ThisAddIn.Application.Workbooks)
            {
                if (wb != workbook)
                {
                    wbs.Add((Excel.Workbook)wb);
                }
            }

            if (wbs.Count == 0)
            {
                wbstates.Clear();
                SetUIStateNoWorkbooks();
            }
        }

        private void WorkbookClose(Excel.Workbook workbook, ref bool Cancel)
        {
            wbstates.Remove(workbook);
            if (wbstates.Count == 0)
            {
                SetUIStateNoWorkbooks();
            }
        }

        #region BUTTON_HANDLERS
        private void AnalyzeButton_Click(object sender, RibbonControlEventArgs e)
        {
            // check for debug easter egg
            if ((System.Windows.Input.Keyboard.Modifiers & System.Windows.Input.ModifierKeys.Alt) > 0)
            {
                current_workbook.DebugMode = true;
            }

            var sig = GetSignificance(this.SensitivityTextBox.Text, this.SensitivityTextBox.Label);
            if (sig == FSharpOption<double>.None)
            {
                return;
            }
            else
            {
                current_workbook.ToolSignificance = sig.Value;
                try
                {
                    current_workbook.Analyze(WorkbookState.MAX_DURATION_IN_MS);
                    current_workbook.Flag();
                    SetUIState(current_workbook);
                }
                catch (Parcel.ParseException ex)
                {
                    System.Windows.Forms.Clipboard.SetText(ex.Message);
                    System.Windows.Forms.MessageBox.Show("Could not parse the formula string:\n" + ex.Message);
                    return;
                }
                catch (System.OutOfMemoryException ex)
                {
                    System.Windows.Forms.MessageBox.Show("Insufficient memory to perform analysis.");
                    return;
                }
            }
        }

        private void FixErrorButton_Click(object sender, RibbonControlEventArgs e)
        {
            current_workbook.FixError(SetUIState);
        }

        private void MarkAsOKButton_Click(object sender, RibbonControlEventArgs e)
        {
            current_workbook.MarkAsOK();
            SetUIState(current_workbook);
        }

        private void StartOverButton_Click(object sender, RibbonControlEventArgs e)
        {
            current_workbook.ResetTool();
            SetUIState(current_workbook);
        }

        private void ToDOTButton_Click(object sender, RibbonControlEventArgs e)
        {
            var app = Globals.ThisAddIn.Application;
            var dag = new DAG(app.ActiveWorkbook, app, true);
            System.Windows.Forms.Clipboard.SetText(dag.ToDOT());
            System.Windows.Forms.MessageBox.Show("In clipboard");
        }

        private void AboutCheckCell_Click(object sender, RibbonControlEventArgs e)
        {
            var ab = new AboutBox();
            ab.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            ab.Show();
        }
        #endregion BUTTON_HANDLERS

        #region UTILITY_FUNCTIONS
        private static FSharpOption<double> GetSignificance(string input, string label)
        {
            var errormsg = label + " must be a value between 0 and 100";
            var significance = 0.95;

            try
            {
                significance = (100.0 - Double.Parse(input)) / 100.0;
            }
            catch
            {
                System.Windows.Forms.MessageBox.Show(errormsg);
            }

            if (significance < 0 || significance > 100)
            {
                System.Windows.Forms.MessageBox.Show(errormsg);
            }

            return FSharpOption<double>.Some(significance);
        }
        #endregion UTILITY_FUNCTIONS
    }
}
