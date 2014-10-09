using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using DataDebugMethods;

namespace DataDebug
{
    public class WorkbookState
    {
        #region CONSTANTS
        // e * 1000
        //public readonly static int NBOOTS = (int)(Math.Ceiling(1000 * Math.Exp(1.0)));
        public readonly static int NBOOTS = 10;
        public readonly static long MAX_DURATION_IN_MS = 5L * 60L * 1000L;  // 5 minutes
        public readonly static System.Drawing.Color GREEN = System.Drawing.Color.Green;
        public readonly static bool IGNORE_PARSE_ERRORS = true;
        #endregion CONSTANTS

        private Excel.Application _app;
        private Excel.Workbook _workbook;
        private double _tool_significance = 0.95;
        private List<RibbonHelper.CellColor> _colors;
        private HashSet<AST.Address> _tool_highlights = new HashSet<AST.Address>();
        private HashSet<AST.Address> _output_highlights = new HashSet<AST.Address>();
        private HashSet<AST.Address> _known_good = new HashSet<AST.Address>();
        private IEnumerable<KeyValuePair<AST.Address, int>> _flaggable;
        private AST.Address _flagged_cell;
        private DAG dag;

        #region BUTTON_STATE
        private bool _button_Analyze_enabled = true;
        private bool _button_MarkAsOK_enabled = false;
        private bool _button_FixError_enabled = false;
        private bool _button_clearColoringButton_enabled = false;
        #endregion BUTTON_STATE

        public WorkbookState(Excel.Application app, Excel.Workbook workbook)
        {
            _app = app;
            _workbook = workbook;
            _colors = RibbonHelper.SaveColors2(workbook);
        }

        public double ToolSignificance
        {
            get { return _tool_significance; }
            set { _tool_significance = value; }
        }

        public bool Analyze_Enabled
        {
            get { return _button_Analyze_enabled; }
            set { _button_Analyze_enabled = value; }
        }

        public bool MarkAsOK_Enabled
        {
            get { return _button_MarkAsOK_enabled; }
            set { _button_MarkAsOK_enabled = value; }
        }

        public bool FixError_Enabled
        {
            get { return _button_FixError_enabled; }
            set { _button_FixError_enabled = value; }
        }
        public bool ClearColoringButton_Enabled
        {
            get { return _button_clearColoringButton_enabled; }
            set { _button_clearColoringButton_enabled = value; }
        }

        public void Analyze(long max_duration_in_ms)
        {
            var sw = new System.Diagnostics.Stopwatch();
            sw.Start();

            using (var pb = new ProgBar())
            {
                // Disable screen updating during analysis to speed things up
                _app.ScreenUpdating = false;

                // Build dependency graph (modifies data)
                try
                {
                    dag = new DAG(_app.ActiveWorkbook, _app, IGNORE_PARSE_ERRORS);
                }
                catch (ExcelParserUtility.ParseException e)
                {
                    // cleanup UI and then rethrow
                    _app.ScreenUpdating = true;
                    throw e;
                }

                if (dag.terminalInputVectors().Length == 0)
                {
                    System.Windows.Forms.MessageBox.Show("This spreadsheet contains no functions that take inputs.");
                    _app.ScreenUpdating = true;
                    _flaggable = new KeyValuePair<AST.Address,int>[0];
                    return;
                }

                // Get bootstraps
                var scores = Analysis.DataDebug(NBOOTS, dag, _app, true, true, max_duration_in_ms, sw, _tool_significance, pb)
                                     .OrderByDescending(pair => pair.Value).ToArray();

                int start_ptr = 0;
                int end_ptr = 0;
                List<KeyValuePair<AST.Address, int>> high_scores = new List<KeyValuePair<AST.Address, int>>();

                while ((double)start_ptr / scores.Length < 1.0 - _tool_significance) //the start of this score region is before the cutoff
                {
                    //while the scores at the start and end pointers are the same, bump the end pointer
                    while (end_ptr < scores.Length && scores[start_ptr].Value == scores[end_ptr].Value)
                    {
                        end_ptr++;
                    }
                    //Now the end_pointer points to the first index with a lower score
                    //If the number of entries with the current value is fewer than the significance cutoff, add all values of this score to the high_scores list; the number of entries is equal to the end_ptr since end_ptr is zero-based
                    //There is some added "wiggle room" to the cutoff, so that the last entry is allowed to straddle the cutoff bound.
                    //  To do this, we add (1 / total number of entries) to the cutoff
                    //The purpose of the wiggle room is to allow us to deal with small ranges (less than 20 entries), since a single entry accounts
                    //for more than 5% of the total.
                    //      Note: tool_significance is along the lines of 0.95 (not 0.05).
                    if ((double)end_ptr / scores.Length < 1.0 - _tool_significance + (double)1.0 / scores.Length)
                    {
                        //add all values of the current score to high_scores list
                        for (; start_ptr < end_ptr; start_ptr++)
                        {
                            high_scores.Add(scores[start_ptr]);
                        }
                        //Increment the start pointer to the start of the next score region
                        start_ptr++;
                    }
                    else    //if this score region extends past the cutoff, we don't add any of its values to the high_scores list, and stop
                    {
                        break;
                    }
                }

                // filter out cells marked as OK
                _flaggable = high_scores.ToArray().Where(kvp => !_known_good.Contains(kvp.Key));
                
                // Enable screen updating when we're done
                _app.ScreenUpdating = true;

                sw.Stop();
            }
        }

        private void ActivateAndCenterOn(AST.Address cell, Excel.Application app)
        {
            // go to worksheet
            RibbonHelper.GetWorksheetByName(cell.A1Worksheet(), _workbook.Worksheets).Activate();

            // COM object
            var comobj = cell.GetCOMObject(app);

            // center screen on cell
            var visible_columns = app.ActiveWindow.VisibleRange.Columns.Count;
            var visible_rows = app.ActiveWindow.VisibleRange.Rows.Count;
            app.Goto(comobj, true);
            app.ActiveWindow.SmallScroll(Type.Missing, visible_rows / 2, Type.Missing, visible_columns / 2);

            // select highlighted cell
            // center on highlighted cell
            comobj.Select();

        }

        public void Flag()
        {
            //filter known_good
            _flaggable = _flaggable.Where(kvp => !_known_good.Contains(kvp.Key));
            if (_flaggable.Count() != 0)
            {
                // get TreeNode corresponding to most unusual score
                _flagged_cell = _flaggable.First().Key;
            }
            else
            {
                _flagged_cell = null;
            }

            if (_flagged_cell == null)
            {
                System.Windows.Forms.MessageBox.Show("No bugs remain.");
                ResetTool();
            }
            else
            {

                // TODO: test after AEC; problematic when highlighted value is not data
                //TreeNode flagged_node;
                //if (data.cell_nodes.TryGetValue(flagged_cell, out flagged_node))
                //{
                //    // only do the following if the cell is a data cell
                //    if (flagged_node.hasOutputs())
                //    {
                //        foreach (var output in flagged_node.getOutputs())
                //        {
                //            exploreNode(output);
                //        }
                //    }
                //}

                _flagged_cell.GetCOMObject(_app).Interior.Color = System.Drawing.Color.Red;
                _tool_highlights.Add(_flagged_cell);

                // go to highlighted cell
                ActivateAndCenterOn(_flagged_cell, _app);

                // enable auditing buttons
                SetTool(active: true);
            }
        }

        private void RestoreOutputColors()
        {
            if (_workbook != null)
            {
                RibbonHelper.RestoreColors2(_colors, _output_highlights);
            }
            _output_highlights.Clear();
        }

        public void ResetTool()
        {
            if (_workbook != null)
            {
                RibbonHelper.RestoreColors2(_colors, _tool_highlights);
            }

            _known_good.Clear();
            _tool_highlights.Clear();
            SetTool(active: false);
        }

        private void SetTool(bool active)
        {
            _button_MarkAsOK_enabled = active;
            _button_FixError_enabled = active;
            _button_clearColoringButton_enabled = active;
            _button_Analyze_enabled = !active;
        }

        private static void RunSimulations(Excel.Application app, Excel.Workbook wb, Random rng, UserSimulation.Classification c, string output_dir, double thresh, ProgBar pb)
        {
            // number of bootstraps
            var NBOOTS = 2700;

            // the full path of this workbook
            var filename = app.ActiveWorkbook.Name;

            // the default output filename
            var r = new System.Text.RegularExpressions.Regex(@"(.+)\.xls|xlsx", System.Text.RegularExpressions.RegexOptions.Compiled);
            var default_output_file = "simulation_results.csv";
            var default_log_file = r.Match(filename).Groups[1].Value + ".iterlog.csv";

            // save file location (will append for additional runs)
            var savefile = System.IO.Path.Combine(output_dir, default_output_file);

            // log file location (new file for each new workbook)
            var logfile = System.IO.Path.Combine(output_dir, default_log_file);

            // disable screen updating
            app.ScreenUpdating = false;

            // run simulations
            UserSimulation.Config.RunSimulationPaperMain(app, wb, NBOOTS, 0.95, thresh, c, rng, savefile, MAX_DURATION_IN_MS, logfile, pb, IGNORE_PARSE_ERRORS);

            // enable screen updating
            app.ScreenUpdating = true;
        }

        private static void RunProportionExperiment(Excel.Application app, Excel.Workbook wb, Random rng, UserSimulation.Classification c, string output_dir, double thresh, ProgBar pb)
        {
            // number of bootstraps
            var NBOOTS = 2700;

            // the full path of this workbook
            var filename = app.ActiveWorkbook.Name;

            // the default output filename
            var r = new System.Text.RegularExpressions.Regex(@"(.+)\.xls|xlsx", System.Text.RegularExpressions.RegexOptions.Compiled);
            var default_output_file = "simulation_results.csv";
            var default_log_file = r.Match(filename).Groups[1].Value + ".iterlog.csv";

            // save file location (will append for additional runs)
            var savefile = System.IO.Path.Combine(output_dir, default_output_file);

            // log file location (new file for each new workbook)
            var logfile = System.IO.Path.Combine(output_dir, default_log_file);

            // disable screen updating
            app.ScreenUpdating = false;

            // run simulations
            UserSimulation.Config.RunProportionExperiment(app, wb, NBOOTS, 0.95, thresh, c, rng, savefile, MAX_DURATION_IN_MS, logfile, pb, IGNORE_PARSE_ERRORS);

            // enable screen updating
            app.ScreenUpdating = true;
        }

        private static void RunSubletyExperiment(Excel.Application app, Excel.Workbook wb, Random rng, UserSimulation.Classification c, string output_dir, double thresh, ProgBar pb)
        {
            // number of bootstraps
            var NBOOTS = 2700;

            // the full path of this workbook
            var filename = app.ActiveWorkbook.Name;

            // the default output filename
            var r = new System.Text.RegularExpressions.Regex(@"(.+)\.xls|xlsx", System.Text.RegularExpressions.RegexOptions.Compiled);
            var default_output_file = "simulation_results.csv";
            var default_log_file = r.Match(filename).Groups[1].Value + ".iterlog.csv";

            // save file location (will append for additional runs)
            var savefile = System.IO.Path.Combine(output_dir, default_output_file);

            // log file location (new file for each new workbook)
            var logfile = System.IO.Path.Combine(output_dir, default_log_file);

            // disable screen updating
            app.ScreenUpdating = false;

            // run simulations
            if (!UserSimulation.Config.RunSubletyExperiment(app, wb, NBOOTS, 0.95, thresh, c, rng, savefile, MAX_DURATION_IN_MS, logfile, pb, IGNORE_PARSE_ERRORS))
            {
                System.Windows.Forms.MessageBox.Show("This spreadsheet contains no numeric inputs.");
            }

            // enable screen updating
            app.ScreenUpdating = true;
        }

        internal void MarkAsOK()
        {
            // the user told us that the cell was OK
            _known_good.Add(_flagged_cell);

            // set the color of the cell to green
            var cell = _flagged_cell.GetCOMObject(_app);
            cell.Interior.Color = GREEN;

            // restore output colors
            RestoreOutputColors();

            // flag another value
            Flag();
        }

        internal void FixError(Action<WorkbookState> setUIState)
        {
            var cell = _flagged_cell.GetCOMObject(_app);
            // this callback gets run when the user clicks "OK"
            System.Action callback = () =>
            {
                // add the cell to the known good list
                _known_good.Add(_flagged_cell);

                // unflag the cell
                _flagged_cell = null;
                try
                {
                    // when a user fixes something, we need to re-run the analysis
                    Analyze(MAX_DURATION_IN_MS);
                    // and flag again
                    Flag();
                    // and then set the UI state
                    setUIState(this);
                }
                catch (ExcelParserUtility.ParseException ex)
                {
                    System.Windows.Forms.Clipboard.SetText(ex.Message);
                    System.Windows.Forms.MessageBox.Show("Could not parse the formula string:\n" + ex.Message);
                    return;
                }
            };
            // show the form
            var fixform = new CellFixForm(cell, GREEN, callback);
            fixform.Show();

            // restore output colors
            RestoreOutputColors();
        }
    }
}
