﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Diagnostics;
using Excel = Microsoft.Office.Interop.Excel;
using DataDebugMethods;
using Depends;
using CellDict = System.Collections.Generic.Dictionary<AST.Address, string>;
using TreeScore = System.Collections.Generic.Dictionary<AST.Address, int>;
using ErrorDict = System.Collections.Generic.Dictionary<AST.Address, double>;

namespace UserSimulation
{
    public static class SimulationStep
    {
        // this function returns an address but also updates
        // the filtered_high_scores list
        public static AST.Address CheckCell_Step(UserResults o,
                                           double significance,
                                           CutoffKind ck,
                                           int nboots,
                                           DAG dag,
                                           Excel.Application app,
                                           bool weighted,
                                           bool all_outputs,
                                           bool run_bootstrap,
                                           HashSet<AST.Address> known_good,
                                           ref List<KeyValuePair<AST.Address, int>> filtered_high_scores,
                                           long max_duration_in_ms,
                                           Stopwatch sw,
                                           ProgBar pb)
        {
            // Get bootstraps
            // The bootstrap should only re-run if there is a correction made, 
            //      not when something is marked as OK (isn't one of the introduced errors)
            // The list of suspected cells doesn't change when we mark something as OK,
            //      we just move on to the next thing in the list
            if (run_bootstrap)
            {
                TreeScore scores = Analysis.DataDebug(nboots, dag, app, weighted, all_outputs, max_duration_in_ms, sw, significance, pb);

                // apply a threshold to the scores
                filtered_high_scores = ck.applyCutoff(scores, known_good);
            }
            else  //if no corrections were made (a cell was marked as OK, not corrected)
            {
                //re-filter out cells marked as OK
                filtered_high_scores = filtered_high_scores.Where(kvp => !known_good.Contains(kvp.Key)).ToList();
            }

            if (filtered_high_scores.Count() != 0)
            {
                // get AST.Address corresponding to most unusual score
                return filtered_high_scores[0].Key;
            }
            else
            {
                return null;
            }
        }

        public static AST.Address NormalPerRange_Step(DAG dag,
                                                      Excel.Workbook wb,
                                                      HashSet<AST.Address> known_good,
                                                      long max_duration_in_ms,
                                                      Stopwatch sw)
        {
            AST.Address flagged_cell = null;

            //Generate normal distributions for every input range until an error is found
            //Then break out of the loop and report it.
            foreach (var vect_addr in dag.allVectors())
            {
                var normal_dist = new DataDebugMethods.NormalDistribution(dag.getCOMRefForRange(vect_addr).Range);

                // Get top outlier which has not been inspected already
                if (normal_dist.getErrorsCount() > 0)
                {
                    for (int i = 0; i < normal_dist.getErrorsCount(); i++)
                    {
                        // check for timeout
                        if (sw.ElapsedMilliseconds > max_duration_in_ms)
                        {
                            throw new TimeoutException("Timeout exception in NormalPerRange_Step.");
                        }

                        var flagged_com = normal_dist.getErrorAtPosition(i);
                        flagged_cell = ParcelCOMShim.Address.AddressFromCOMObject(flagged_com, wb);
                        if (known_good.Contains(flagged_cell))
                        {
                            flagged_cell = null;
                        }
                        else
                        {
                            break;
                        }
                    }
                }
                //If a cell is flagged, do not move on to the next range (if you do, you'll overwrite the flagged_cell
                if (flagged_cell != null)
                {
                    break;
                }
            }

            return flagged_cell;
        }

        public static AST.Address NormalAllOutputs_Step(DAG dag,
                                                        Excel.Application app,
                                                        Excel.Workbook wb,
                                                        HashSet<AST.Address> known_good,
                                                        long max_duration_in_ms,
                                                        Stopwatch sw)
        {
            AST.Address flagged_cell = null;

            //Generate a normal distribution for the entire set of inputs
            var normal_dist = new DataDebugMethods.NormalDistribution(dag.terminalInputVectors(), app);

            // Get top outlier
            if (normal_dist.getErrorsCount() > 0)
            {
                for (int i = 0; i < normal_dist.getErrorsCount(); i++)
                {
                    // check for timeout
                    if (sw.ElapsedMilliseconds > max_duration_in_ms)
                    {
                        throw new TimeoutException("Timeout exception in NormalAllOutputs_Step.");
                    }

                    var flagged_com = normal_dist.getErrorAtPosition(i);
                    flagged_cell = ParcelCOMShim.Address.AddressFromCOMObject(flagged_com, wb);
                    if (known_good.Contains(flagged_cell))
                    {
                        flagged_cell = null;
                    }
                    else
                    {
                        break;
                    }
                }
            }

            return flagged_cell;
        }
    }
}
