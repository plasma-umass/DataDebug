using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
using System.Text.RegularExpressions;
using Microsoft.FSharp.Core;

namespace DataDebugMethods
{
    public static class DependenceAnalysis
    {
        public static DAG constructDAG(Excel.Workbook wb, Excel.Application app, bool ignore_parse_errors)
        {
            return constructDAG(wb, app, null, ignore_parse_errors);
        }

        // This method constructs the dependency graph from the workbook.
        public static DAG constructDAG(Excel.Workbook wb, Excel.Application app, ProgBar pb, bool ignore_parse_errors)
        {
            // use a fast array read to find all cell & formula addresses
            var dag = new DAG(wb, app);

            // extract references from formulas
            foreach(AST.Address formula_addr in dag.GetFormulaAddrs())
            {
                // get COMRef read earlier
                var formula_ref = dag.GetCOMRefForAddress(formula_addr);

                foreach (AST.Range vector_rng in ExcelParserUtility.GetRangeReferencesFromFormula(formula_ref, ignore_parse_errors))
                {
                    // fetch/create COMRef, as appropriate
                    var vector_ref = dag.MakeInputVectorCOMRef(vector_rng);

                    // link formula and input vector
                    dag.LinkInputVector(formula_addr, vector_rng);

                    foreach (AST.Address input_single in vector_rng.Addresses()) {
                        // link input vector and single input
                        
                    }

                    // if num single inputs = num formulas,
                    // mark vector as non-perturbable
                }

                foreach (AST.Address input_single in ExcelParserUtility.GetSingleCellReferencesFromFormula(formula_ref, ignore_parse_errors))
                {
                    // link formula and single input
                }
            }

            return dag;
        }

       

    } // ConstructTree class ends here
} // namespace ends here
