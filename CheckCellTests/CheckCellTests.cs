using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using DataDebugMethods;

namespace CheckCellTests
{
    [TestClass]
    public class CheckCellTests
    {
        [TestMethod]
        public void TestGetFormulaRanges()
        {
            // worksheet indices; watch out! the second index here is the NUMBER of elements, NOT the max value!
            var e = Enumerable.Range(1,10);

            // rnd, for random formulae assignment
            Random rand = new Random();

            // new Excel instance
            Excel.Application app = new Excel.Application();

            // create new workbook
            Excel.Workbook wb = app.Workbooks.Add();

            // get a reference to the worksheet array
            // By default, workbooks have three blank worksheets.
            Excel.Sheets ws = wb.Worksheets;
            
            // add some worksheets
            foreach (int i in e)
            {
                ws.Add(Missing.Value, Missing.Value, Missing.Value, Missing.Value);
            }

            // gin up some formulae
            Tuple<string,string>[] fs = {new Tuple<string,string>("B4", "=COUNT(A1:A5)"),
                                         new Tuple<string,string>("A6", "=SUM(B5:B40)"),
                                         new Tuple<string,string>("Z2", "=AVERAGE(A1:E1)"),
                                         new Tuple<string,string>("B44", "=MEDIAN(D4:D9)")};

            // to keep track of what we did
            var d = new Dictionary<int,List<Tuple<string,string>>>();

            // add the formulae to the worksheets, randomly
            foreach (int i in e)
            {
                // convert array index into worksheet reference, because
                // GetFormulaRanges returns an array indexed not by formula reference
                // but by the worksheet's index in the global worksheet array
                Excel.Worksheet w = ws[i + 3];

                // init list for each worksheet
                d[i] = new List<Tuple<string,string>>();

                // add the formulae, randomly
                foreach (var f in fs)
                {
                    if (rand.Next(0, 2) == 0)
                    {
                        w.Range[f.Item1, f.Item1].Formula = f.Item2;
                        // keep track of what we did
                        d[i].Add(f);
                    }
                }
            }

            // get the formulae
            Excel.Range[] fs_rs = ConstructTree.GetFormulaRanges(ws, app);

            // there should be e.Count + 3 entries
            // don't forget: workbooks have 3 blank worksheets by default
            if (fs_rs.Length != e.Count() + 3) {
                throw new Exception("ConstructTree.GetFormulaRanges() should return " + e.Count().ToString() + " elements.");
            }

            // make sure that each worksheet's range has the formulae that it should
            bool all_ok = true;
            foreach (int i in e)
            {
                // get a reference to the range
                // note that the array returned by ConstructTree.GetFormulaRanges is zero-based unlike Excel's Workbook.Worksheets list.
                // Thus we only adjust i by 2.
                Excel.Range r = fs_rs[i + 2];

                // check that all formulae for this worksheet are accounted for
                bool r_ok = d[i].Aggregate(true, (bool acc, Tuple<string,string> f) => {
                                bool found = false;
                                foreach(Excel.Range cell in r) {
                                    if (cell.Formula == f.Item2) {
                                        found = true;
                                    }
                                }
                                return acc && found;
                            });

                all_ok = all_ok && r_ok;
            }

            if (!all_ok) {
                throw new Exception("ConstructTree.GetFormulaRanges() failed to return all of the formulae that were added.");
            }
        }
    }
}
