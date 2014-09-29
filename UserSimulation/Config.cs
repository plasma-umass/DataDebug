using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DataDebugMethods;
using Excel = Microsoft.Office.Interop.Excel;
using TreeNode = DataDebugMethods.TreeNode;
using CellDict = System.Collections.Generic.Dictionary<AST.Address, string>;
using TreeScore = System.Collections.Generic.Dictionary<DataDebugMethods.TreeNode, int>;
using ErrorDict = System.Collections.Generic.Dictionary<AST.Address, double>;

namespace UserSimulation
{
    public static class Config
    {
        public static void RunSimulationPaperMain(Excel.Application app, Excel.Workbook wbh, int nboots, double significance, double threshold, UserSimulation.Classification c, Random r, String outfile, long max_duration_in_ms, String logfile, ProgBar pb, bool ignore_parse_errors)
        {
            // record intitial state of spreadsheet
            var prepdata = Prep.PrepSimulation(app, wbh, pb, ignore_parse_errors);

            // generate errors
            CellDict errors = UserSimulation.Utility.GenImportantErrors(prepdata.terminal_formula_nodes,
                                                               prepdata.original_inputs,
                                                               5,
                                                               prepdata.correct_outputs,
                                                               app,
                                                               wbh,
                                                               c);
            // run paper simulations
            RunSimulation(app, wbh, nboots, significance, threshold, c, r, outfile, max_duration_in_ms, logfile, pb, prepdata, errors);
        }

        public static void RunProportionExperiment(Excel.Application app, Excel.Workbook wbh, int nboots, double significance, double threshold, UserSimulation.Classification c, Random r, String outfile, long max_duration_in_ms, String logfile, ProgBar pb, bool ignore_parse_errors)
        {
            // record intitial state of spreadsheet
            var prepdata = Prep.PrepSimulation(app, wbh, pb, ignore_parse_errors);

            // init error generator
            var eg = new ErrorGenerator();

            // get inputs as an array of addresses to facilitate random selection
            // DATA INPUTS ONLY
            var inputs = prepdata.graph.TerminalInputCells().Select(n => n.GetAddress()).ToArray<AST.Address>();

            // sanity check: all of the inputs should also be in prepdata.original_inputs
            foreach (AST.Address addr in inputs)
            {
                if (!prepdata.original_inputs.ContainsKey(addr))
                {
                    throw new Exception("Missing address!");
                }
            }

            for (int i = 0; i < 100; i++)
            {
                // randomly choose an input address
                AST.Address rand_addr = inputs[r.Next(inputs.Length)];

                // get the value
                String input_value = prepdata.original_inputs[rand_addr];

                // perturb it
                String erroneous_input = eg.GenerateErrorString(input_value, c);

                // create an error dictionary with this one perturbed value
                var errors = new CellDict();
                errors.Add(rand_addr, erroneous_input);

                // run simulations; simulation code does insertion of errors and restore of originals
                RunSimulation(app, wbh, nboots, significance, threshold, c, r, outfile, max_duration_in_ms, logfile, pb, prepdata, errors);
            }
        }

        public static bool RunSubletyExperiment(Excel.Application app, Excel.Workbook wbh, int nboots, double significance, double threshold, UserSimulation.Classification c, Random r, String outfile, long max_duration_in_ms, String logfile, ProgBar pb, bool ignore_parse_errors)
        {
            // record intitial state of spreadsheet
            var prepdata = Prep.PrepSimulation(app, wbh, pb, ignore_parse_errors);

            // init error generator
            var eg = new ErrorGenerator();

            // get inputs as an array of addresses to facilitate random selection
            // DATA INPUTS ONLY
            var inputs = prepdata.graph.TerminalInputCells().Select(n => n.GetAddress()).ToArray<AST.Address>();

            for (int i = 0; i < 100; i++)
            {
                // randomly choose a *numeric* input
                // TODO: use Fischer-Yates and take values until
                // either we have a satisfactory input value or none
                // remain
                var rnd_addrs = inputs.Shuffle().ToList();
                bool num_found = false;
                String input_string;
                double input_value;
                AST.Address rand_addr;
                do
                {
                    // randomly choose an address; if there are none left, fail
                    if (rnd_addrs.Count == 0) {
                        return false;
                    }
                    rand_addr = rnd_addrs.First();
                    rnd_addrs = rnd_addrs.Skip(1).ToList();

                    // get the value
                    input_string = prepdata.original_inputs[rand_addr];

                    // try parsing it
                    if (Double.TryParse(input_string, out input_value))
                    {
                        num_found = true;
                    }
                } while (!num_found);

                // perturb it
                String erroneous_input = eg.GenerateSubtleErrorString(input_value, c);

                // create an error dictionary with this one perturbed value
                var errors = new CellDict();
                errors.Add(rand_addr, erroneous_input);

                // run simulations; simulation code does insertion of errors and restore of originals
                RunSimulation(app, wbh, nboots, significance, threshold, c, r, outfile, max_duration_in_ms, logfile, pb, prepdata, errors);
            }

            return true;
        }

        public static void RunSimulation(Excel.Application app, Excel.Workbook wbh, int nboots, double significance, double threshold, UserSimulation.Classification c, Random r, String outfile, long max_duration_in_ms, String logfile, ProgBar pb, PrepData prepdata, CellDict errors)
        {
            pb.IncrementProgress(16);

            // write header if needed
            if (!System.IO.File.Exists(outfile))
            {
                System.IO.File.AppendAllText(outfile, Simulation.HeaderRowForCSV());
            }

            // CheckCell weighted, all outputs, quantile
            //var s_1 = new UserSimulation.Simulation();
            //s_1.RunFromBatch(nboots,                                   // number of bootstraps
            //                    wbh.FullName,                          // Excel filename
            //                    significance,                          // statistical significance threshold for hypothesis test
            //                    app,                                   // Excel.Application
            //                    new QuantileCutoff(0.05),              // max % extreme values to flag
            //                    c,                                     // classification data
            //                    r,                                     // random number generator
            //                    UserSimulation.AnalysisType.CheckCell5,// analysis type
            //                    true,                                  // weighted analysis
            //                    true,                                  // use all outputs for analysis
            //                    prepdata.graph,                                 // AnalysisData
            //                    wbh,                                   // Excel.Workbook
            //                    errors,                                // pre-generated errors
            //                    prepdata.terminal_input_nodes,                  // input range nodes
            //                    prepdata.terminal_formula_nodes,                // output nodes
            //                    prepdata.original_inputs,                       // original input values
            //                    prepdata.correct_outputs,                       // original output values
            //                    max_duration_in_ms,                    // max duration of simulation 
            //                    logfile);
            //System.IO.File.AppendAllText(outfile, s_1.FormatResultsAsCSV());
            //pb.IncrementProgress(16);

            // CheckCell weighted, all outputs, quantile
            var s_4 = new UserSimulation.Simulation();
            s_4.RunFromBatch(nboots,                                   // number of bootstraps
                                wbh.FullName,                          // Excel filename
                                significance,                          // statistical significance of threshold
                                app,                                   // Excel.Application
                                new QuantileCutoff(0.10),              // max % extreme values to flag
                                c,                                     // classification data
                                r,                                     // random number generator
                                UserSimulation.AnalysisType.CheckCell10,// analysis type
                                true,                                  // weighted analysis
                                true,                                  // use all outputs for analysis
                                prepdata.graph,                                 // AnalysisData
                                wbh,                                   // Excel.Workbook
                                errors,                                // pre-generated errors
                                prepdata.terminal_input_nodes,                  // input range nodes
                                prepdata.terminal_formula_nodes,                // output nodes
                                prepdata.original_inputs,                       // original input values
                                prepdata.correct_outputs,                       // original output values
                                max_duration_in_ms,                    // max duration of simulation 
                                logfile);
            System.IO.File.AppendAllText(outfile, s_4.FormatResultsAsCSV());
            pb.IncrementProgress(16);

            // Normal, all inputs
            var s_2 = new UserSimulation.Simulation();
            s_2.RunFromBatch(nboots,                                   // irrelevant
                                wbh.FullName,                              // Excel filename
                                significance,                          // normal cutoff?
                                app,                                   // Excel.Application
                                new NormalCutoff(threshold),           // ??
                                c,                                     // classification data
                                r,                                     // random number generator
                                UserSimulation.AnalysisType.NormalAllInputs,   // analysis type
                                true,                                  // irrelevant
                                true,                                  // irrelevant
                                prepdata.graph,                                 // AnalysisData
                                wbh,                                   // Excel.Workbook
                                errors,                                // pre-generated errors
                                prepdata.terminal_input_nodes,                  // input range nodes
                                prepdata.terminal_formula_nodes,                // output nodes
                                prepdata.original_inputs,                       // original input values
                                prepdata.correct_outputs,                       // original output values
                                max_duration_in_ms,                    // max duration of simulation 
                                logfile);
            System.IO.File.AppendAllText(outfile, s_2.FormatResultsAsCSV());
            pb.IncrementProgress(16);

            // Normal, range inputs
            //var s_3 = new UserSimulation.Simulation();
            //s_3.RunFromBatch(nboots,                                   // irrelevant
            //                    wbh.FullName,                              // Excel filename
            //                    significance,                          // normal cutoff?
            //                    app,                                   // Excel.Application
            //                    new NormalCutoff(threshold),           // ??
            //                    c,                                     // classification data
            //                    r,                                     // random number generator
            //                    UserSimulation.AnalysisType.NormalPerRange,   // analysis type
            //                    true,                                  // irrelevant
            //                    true,                                  // irrelevant
            //                    prepdata.graph,                                 // AnalysisData
            //                    wbh,                                   // Excel.Workbook
            //                    errors,                                // pre-generated errors
            //                    prepdata.terminal_input_nodes,                  // input range nodes
            //                    prepdata.terminal_formula_nodes,                // output nodes
            //                    prepdata.original_inputs,                       // original input values
            //                    prepdata.correct_outputs,                       // original output values
            //                    max_duration_in_ms,                    // max duration of simulation 
            //                    logfile);
            //System.IO.File.AppendAllText(outfile, s_3.FormatResultsAsCSV());
            //pb.IncrementProgress(20);
        }
    }
}
