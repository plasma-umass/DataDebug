using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DataDebugMethods;
using Excel = Microsoft.Office.Interop.Excel;
using ColorDict = System.Collections.Generic.Dictionary<Microsoft.Office.Interop.Excel.Workbook, System.Collections.Generic.List<AST.Address>>;
using OptString = Microsoft.FSharp.Core.FSharpOption<string>;
using OptTuple = Microsoft.FSharp.Core.FSharpOption<System.Tuple<UserSimulation.Classification, string>>;

namespace DataDebug
{
    static class RibbonHelper
    {
        private static int TRANSPARENT_COLOR_INDEX = -4142;  //-4142 is the transparent default background

        public static OptTuple getExperimentInputs()
        {
            // ask the user for the classification data
            var ofd = new System.Windows.Forms.OpenFileDialog();
            ofd.ShowHelp = true;
            ofd.FileName = "ClassificationData-2013-11-14.bin";
            ofd.Title = "Please select a classification data input file.";
            if (ofd.ShowDialog() != System.Windows.Forms.DialogResult.OK)
            {
                return OptTuple.None;
            }
            var c = UserSimulation.Classification.Deserialize(ofd.FileName);

            // ask the user where the output data should go
            var cdd = new System.Windows.Forms.FolderBrowserDialog();
            cdd.Description = "Please choose an output folder.";
            if (cdd.ShowDialog() != System.Windows.Forms.DialogResult.OK)
            {
                return OptTuple.None;
            }
            var foo = new Tuple<UserSimulation.Classification, string>(c, cdd.SelectedPath);
            return OptTuple.Some(foo);
        }

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
                return AST.Address.AddressFromCOMObject(_cellCOM, _cellCOM.Application.ActiveWorkbook);
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
