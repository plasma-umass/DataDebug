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
