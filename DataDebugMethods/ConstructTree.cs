using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;

namespace DataDebugMethods
{
    public static class ConstructTree
    {
        public static int CountFormulaCells(Excel.Range[] rs)
        {
            int count = 0;
            foreach (var r in rs)
            {
                count += r.Cells.Count;
            }
            return count;
        }

        public static Excel.Range[] GetFormulaRanges(Excel.Sheets ws, Excel.Application app)
        {
            Excel.Range[] analysisRanges = new Excel.Range[ws.Count]; //This keeps track of the range to be analyzed in every worksheet of the workbook

            int worksheet_index = 0; // keeps track of which worksheet we are currently examining
            foreach (Excel.Worksheet w in ws)
            {
                Excel.Range formula_cells = null;
                // iterate over all of the cells in a particular worksheet
                // these actually are cells, because that's what you get when you
                // iterate over the UsedRange property
                foreach (Excel.Range cell in w.UsedRange)
                {
                    // the cell thinks it has a formula
                    if (cell.HasFormula)
                    {
                        // this is our first time around; formula_cells is not yet set,
                        // so set it
                        if (formula_cells == null)
                        {
                            formula_cells = cell;
                        }
                        // it's not our first time around, so union the current cell with
                        // the previously found formula cell
                        else
                        {
                            formula_cells = app.Union(
                                            cell,
                                            formula_cells,
                                            Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                            Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                            Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                            Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                        }
                    }
                }
                // we found at least one cell
                if (formula_cells != null)
                {
                    analysisRanges[worksheet_index] = formula_cells;
                }
                // we found no cells
                else
                {
                    analysisRanges[worksheet_index] = null;
                }
                // point at the next worksheet in analysisRanges
                worksheet_index++;
            }
            return analysisRanges;
        }
    }
}
