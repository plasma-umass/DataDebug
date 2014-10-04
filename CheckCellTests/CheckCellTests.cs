using System;
using System.Text;
using System.Collections;
using System.Linq;
using System.Reflection;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using DataDebugMethods;
using System.Runtime.InteropServices;

namespace CheckCellTests
{
    [TestClass]
    public class CheckCellTests
    {
        public class MockWorkbook
        {
            Excel.Application app;
            Excel.Workbook wb;
            Excel.Sheets ws;

            public MockWorkbook()
            {
                // worksheet indices; watch out! the second index here is the NUMBER of elements, NOT the max value!
                var e = Enumerable.Range(1, 10);

                // new Excel instance
                app = new Excel.Application();

                // create new workbook
                wb = app.Workbooks.Add();

                // get a reference to the worksheet array
                // By default, workbooks have three blank worksheets.
                ws = wb.Worksheets;

                // add some worksheets
                foreach (int i in e)
                {
                    ws.Add(Missing.Value, Missing.Value, Missing.Value, Missing.Value);
                }
            }
            public Excel.Application GetApplication() { return app; }
            public Excel.Workbook GetWorkbook() { return wb; }
            public Excel.Sheets GetWorksheets() { return ws; }
            public Excel.Worksheet GetWorksheet(int idx) { return (Excel.Worksheet)ws[idx]; }
            public static int TestGetRanges(string formula)
            {
                // mock workbook object
                var mwb = new MockWorkbook();
                Excel.Workbook wb = mwb.GetWorkbook();
                Excel.Worksheet ws = mwb.GetWorksheet(1);

                var ranges = ExcelParserUtility.GetReferencesFromFormula(formula, wb, ws, ignore_parse_errors: false);

                return ranges.Count();
            }
            ~MockWorkbook()
            {
                try
                {
                    wb.Close(false, Type.Missing, Type.Missing);
                    app.Quit();
                    Marshal.ReleaseComObject(ws);
                    Marshal.ReleaseComObject(wb);
                    Marshal.ReleaseComObject(app);
                    ws = null;
                    wb = null;
                    app = null;
                }
                catch
                {
                }
            }
        }

        [TestMethod]
        public void TestGetFormulaRanges()
        {
            var mwb = new MockWorkbook();

            // rnd, for random formulae assignment
            Random rand = new Random();

            // gin up some formulae
            Tuple<string,string>[] fs = {new Tuple<string,string>("B4", "=COUNT(A1:A5)"),
                                         new Tuple<string,string>("A6", "=SUM(B5:B40)"),
                                         new Tuple<string,string>("Z2", "=AVERAGE(A1:E1)"),
                                         new Tuple<string,string>("B44", "=MEDIAN(D4:D9)")};

            // to keep track of what we did
            var d = new System.Collections.Generic.Dictionary<Excel.Worksheet, System.Collections.Generic.List<Tuple<string, string>>>();

            // add the formulae to the worksheets, randomly
            foreach (Excel.Worksheet w in mwb.GetWorksheets())
            {
                // init list for each worksheet
                d[w] = new System.Collections.Generic.List<Tuple<string,string>>();

                // add the formulae, randomly
                foreach (var f in fs)
                {
                    if (rand.Next(0, 2) == 0)
                    {
                        w.Range[f.Item1, f.Item1].Formula = f.Item2;
                        // keep track of what we did
                        d[w].Add(f);
                    }
                }
                // we need at least one formula, so add one if the above procedure did not
                if (d[w].Count() == 0)
                {
                    w.Range[fs[0].Item1, fs[0].Item1].Formula = fs[0].Item2;
                }
            }

            // get the formulae; 1 formula per worksheet
            ArrayList fs_rs = DependenceAnalysis.GetFormulaRanges(mwb.GetWorksheets(), mwb.GetApplication());

            // there should be e.Count + 3 entries
            // don't forget: workbooks have 3 blank worksheets by default
            if (fs_rs.Count != mwb.GetWorksheets().Count) {
                throw new Exception("ConstructTree.GetFormulaRanges() should return " + mwb.GetWorksheets().Count.ToString() + " elements.");
            }

            // make sure that each worksheet's range has the formulae that it should
            bool all_ok = true;
            foreach (Excel.Range r in fs_rs)
            {
                // check that all formulae for this worksheet are accounted for
                bool r_ok = d[r.Worksheet].Aggregate(true, (bool acc, Tuple<string,string> f) => {
                                bool found = false;
                                foreach(Excel.Range cell in r) {
                                    if (String.Equals((string)cell.Formula, f.Item2)) {
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
        } // end test

        [TestMethod]
        public void TestGetRanges1()
        {
            var f = "=A1";
            if (MockWorkbook.TestGetRanges(f) != 0)
            {
                throw new Exception("GetReferencesFromFormula should return no ranges for " + f);
            }
        }

        [TestMethod]
        public void TestGetRanges2()
        {
            var f = "=A1:B3";
            if (MockWorkbook.TestGetRanges(f) != 1)
            {
                throw new Exception("GetReferencesFromFormula should return 1 range for " + f);
            }
        }

        [TestMethod]
        public void TestGetRanges3()
        {
            var f = "=SUM(A1:B3)+AVERAGE(C2:C8)";
            if (MockWorkbook.TestGetRanges(f) != 2)
            {
                throw new Exception("GetReferencesFromFormula should return 2 ranges for " + f);
            }
        }
    }
}
