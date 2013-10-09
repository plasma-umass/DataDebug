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

        public class CellColor
        {
            public string getAddress()
            {
                return _addr;
            }
            private Excel.Worksheet _ws;
            private string _addr;
            private int _colorindex;
            private double _color;
            private Excel.Range _cellCOM;

            public CellColor(Excel.Worksheet ws, Excel.Range cellCOM, string address, int colorindex, double color)
            {
                _ws = ws;
                _addr = address;
                _colorindex = colorindex;
                _color = color;
                _cellCOM = cellCOM;
            }
            public AST.Address GetASTAddr()
            {
                return AST.Address.AddressFromCOMObject(_cellCOM,
                                                        new Microsoft.FSharp.Core.FSharpOption<string>(_ws.Name),
                                                        new Microsoft.FSharp.Core.FSharpOption<string>(_cellCOM.Application.ActiveWorkbook.Name),
                                                        new Microsoft.FSharp.Core.FSharpOption<string>(_cellCOM.Application.ActiveWorkbook.FullName));
            }
            public void Restore(HashSet<AST.Address> tool_highlights)
            {
                if (tool_highlights.Contains(this.GetASTAddr()))
                { //this color was set by us, so reset it
                    if (_colorindex == TRANSPARENT_COLOR_INDEX)
                    {
                        _ws.get_Range(_addr).Interior.ColorIndex = _colorindex;
                    }
                    else
                    {
                        _ws.get_Range(_addr).Interior.Color = _color;
                    }
                }
                else { } //the user set this color after the tool was run, so do not reset it
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
                    _l.Add(new CellColor(ws, cell, cell.Address, cell.Interior.ColorIndex, cell.Interior.Color));
                }
            }
            return _l;
        }

        public static void RestoreColors2(List<CellColor> colors, HashSet<AST.Address> tool_highlights)
        {
            foreach (CellColor c in colors)
            {
                c.Restore(tool_highlights);
            }
        }

        public static Excel.Worksheet GetWorksheetByName(string name, Excel.Sheets sheets)
        {
            foreach (Excel.Worksheet ws in sheets)
            {
                if (ws.Name == name)
                {
                    return ws;
                }
            }
            return null;
        }
    }
}
