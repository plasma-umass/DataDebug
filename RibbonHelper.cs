using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DataDebugMethods;
using Excel = Microsoft.Office.Interop.Excel;
using ColorDict = System.Collections.Generic.Dictionary<Microsoft.Office.Interop.Excel.Workbook, System.Collections.Generic.List<DataDebugMethods.TreeNode>>;

namespace DataDebug
{
    static class RibbonHelper
    {
        private static int TRANSPARENT_COLOR_INDEX = -4142;  //-4142 is the transparent default background

        public static void DisplayGraphvizTree(AnalysisData analysisData)
        {
            string gvstr = ConstructTree.GenerateGraphVizTree(analysisData.formula_nodes);
            Display disp = new Display();
            disp.textBox1.Text = gvstr;
            disp.ShowDialog();
        }

        // Clear saved colors if the workbook matches
        public static void DeleteColorsForWorkbook(ref ColorDict color_storage, Excel.Workbook wb)
        {
            if (color_storage.ContainsKey(wb))
            {
                color_storage.Remove(wb);
            }
        }

        // Save current colors
        public static void SaveColors(ref ColorDict color_storage, Excel.Workbook wb)
        {
            List<TreeNode> ts;
            if (!color_storage.TryGetValue(wb, out ts))
            {
                ts = new List<TreeNode>();
                color_storage.Add(wb, ts);
            }

            foreach (Excel.Worksheet ws in wb.Worksheets)
            {
                foreach (Excel.Range cell in ws.UsedRange)
                {
                    //Create a TreeNode for every cell with the name being the cell's address and set the node's worksheet appropriately
                    TreeNode n = new TreeNode(cell, cell.Address, cell.Worksheet, Globals.ThisAddIn.Application.ActiveWorkbook);
                    n.setOriginalColor(System.Drawing.ColorTranslator.FromOle((int)cell.Interior.Color));
                    ts.Add(n);
                }
            }
        }

        public class CellColor
        {
            private Excel.Worksheet _ws;
            private string _addr;
            private int _colorindex;
            private double _color;

            public CellColor(Excel.Worksheet ws, string address, int colorindex, double color)
            {
                _ws = ws;
                _addr = address;
                _colorindex = colorindex;
                _color = color;
            }
            public void Restore()
            {
                if (_colorindex == TRANSPARENT_COLOR_INDEX)
                {
                    _ws.get_Range(_addr).Interior.ColorIndex = _colorindex;
                }
                else
                {
                    _ws.get_Range(_addr).Interior.Color = _color;
                }
            }
        }

        public static List<CellColor> SaveColors2(Excel.Workbook wb)
        {
            //System.Windows.Forms.MessageBox.Show("Saving colors.");
            var _l = new List<CellColor>();
            foreach (Excel.Worksheet ws in wb.Worksheets)
            {
                foreach (Excel.Range cell in ws.UsedRange)
                {
                    _l.Add(new CellColor(ws, cell.Address, cell.Interior.ColorIndex, cell.Interior.Color));
                }
            }
            return _l;
        }

        public static void RestoreColors2(List<CellColor> colors)
        {
            foreach (CellColor c in colors)
            {
                c.Restore();
            }
        }

        // Restore colors to saved value, if we saved them
        public static void RestoreColorsForWorkbook(ref ColorDict color_storage, Excel.Workbook wb)
        {
            List<TreeNode> ts;
            if (color_storage.TryGetValue(wb, out ts))
            {
                foreach (TreeNode t in ts)
                {
                    if (!t.isChart() && !t.isRange())
                    {
                        if (!t.getOriginalColor().Equals("Color [White]"))
                        {
                            t.getWorksheetObject().get_Range(t.getName()).Interior.Color = t.getOriginalColor();
                        }
                        else
                        {
                            t.getWorksheetObject().get_Range(t.getName()).Interior.ColorIndex = TRANSPARENT_COLOR_INDEX;
                        }
                    }
                }

                color_storage.Remove(wb);
            }
        }

    }
}
