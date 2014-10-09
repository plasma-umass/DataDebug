using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DataDebugMethods;
using Excel = Microsoft.Office.Interop.Excel;
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
            private int _colorindex;
            private double _color;

            public CellColor(int colorindex, double color)
            {
                _colorindex = colorindex;
                _color = color;
            }

            public double Color
            {
                get { return _color; }
            }

            public int ColorIndex
            {
                get { return _colorindex; }
            }
        }

        public static Dictionary<AST.Address, CellColor> SaveColors(Excel.Workbook wb)
        {
            var _d = new Dictionary<AST.Address, CellColor>();

            // get names once
            var wbfullname = wb.FullName;
            var wbname = wb.Name;
            var path = wb.Path;
            var wbname_opt = new Microsoft.FSharp.Core.FSharpOption<String>(wbname);
            var path_opt = new Microsoft.FSharp.Core.FSharpOption<String>(path);

            foreach (Excel.Worksheet worksheet in wb.Worksheets)
            {
                // get used range
                Excel.Range urng = worksheet.UsedRange;

                // get dimensions
                var left = urng.Column;                      // 1-based left-hand y coordinate
                var right = urng.Columns.Count + left - 1;   // 1-based right-hand y coordinate
                var top = urng.Row;                          // 1-based top x coordinate
                var bottom = urng.Rows.Count + top - 1;      // 1-based bottom x coordinate

                // get worksheet name
                var wsname = worksheet.Name;
                var wsname_opt = new Microsoft.FSharp.Core.FSharpOption<String>(wsname);

                // init
                int width = right - left + 1;
                int height = bottom - top + 1;

                // if the used range is a single cell, Excel changes the type
                if (left == right && top == bottom)
                {
                    double color = urng.Interior.Color;
                    int coloridx = urng.Interior.ColorIndex;

                    var addr = AST.Address.NewFromR1C1(top, left, wsname_opt, wbname_opt, path_opt);
                    _d.Add(addr, new CellColor(coloridx, color));
                }
                else
                {
                    // array read of colors
                    // note that this is a 1-based 2D multiarray
                    double[,] colors = urng.Interior.Color;
                    int[,] coloridxs = urng.Interior.ColorIndex;

                    // add to dict
                    for (int c = left; c < left + width; c++)
                    {
                        for (int r = top; r < top + height; r++)
                        {
                                var addr = AST.Address.NewFromR1C1(r + top - 1, c + left - 1, wsname_opt, wbname_opt, path_opt);
                                _d.Add(addr, new CellColor(coloridxs[r, c], colors[r, c]));
                        }
                    }
                }
            }
            return _d;
        }

        public static void RestoreColors(Dictionary<AST.Address, CellColor> colors, HashSet<AST.Address> tool_highlights, Excel.Application app)
        {
            foreach (AST.Address addr in tool_highlights)
            {
                var cc = colors[addr];
                var com = addr.GetCOMObject(app);
                com.Interior.Color = cc.Color;
                com.Interior.ColorIndex = cc.ColorIndex;
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
