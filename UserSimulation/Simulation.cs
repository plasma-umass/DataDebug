using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using TreeNode = DataDebugMethods.TreeNode;
using CellDict = System.Collections.Generic.Dictionary<AST.Address, string>;

namespace UserSimulation
{
    public class Simulation
    {
        //CellDict saved_values = new CellDict();
        //HashSet<AST.Address> tool_highlights = new HashSet<AST.Address>();
        //HashSet<AST.Address> known_good = new HashSet<AST.Address>();
        //IEnumerable<Tuple<double, TreeNode>> analysis_results = null;
        //AST.Address flagged_cell = null;

        // create and run a CheckCell simulation
        public Simulation(string filename, double sensitivity, CellDict errors, Excel.Application app)
        {
            // open spreadsheet
            Excel.Workbook wb = DataDebugMethods.Utility.OpenWorkbook(filename, app);

            // save original spreadsheet state
            CellDict original_values = SaveSpreadsheet(wb);


        }

        // save spreadsheet state to a CellDict
        public CellDict SaveSpreadsheet(Excel.Workbook wb)
        {
            var cd = new CellDict();
            foreach (Excel.Worksheet ws in wb.Worksheets)
            {
                foreach (Excel.Range cell in ws.UsedRange)
                {
                    cd.Add(AST.Address.AddressFromCOMObject(cell, wb), System.Convert.ToString(cell.Value2));
                }
            }
            return cd;
        }
    }
}
