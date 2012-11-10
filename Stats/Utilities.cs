using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;

namespace DataDebug.Stats
{
    class Utilities
    {
        public static Double[] GetCurrentRangeAsArray()
        {
            Excel.Range selectedCell = Globals.ThisAddIn.Application.Selection as Excel.Range; // Instances.getApp().ActiveCell;
            object[,] object2d = selectedCell.CurrentRegion.Value2;
            List<Double> r_out = new List<Double>();

            if (object2d != null)
            {
                foreach (Double cell in object2d)
                {
                    r_out.Add(cell);
                }
            }
            return r_out.ToArray<Double>();
        }

        public static Excel.Range GetCurrentRange()
        {
            Excel.Range selectedCell = Globals.ThisAddIn.Application.Selection as Excel.Range; //Instances.getApp().ActiveCell;
            return selectedCell.CurrentRegion;
        }

        public static void ColorCellListByName(Dictionary<Excel.Range, System.Drawing.Color> cells, String color)
        {
            foreach (KeyValuePair<Excel.Range, System.Drawing.Color> cell in cells)
            {
                cell.Key.Interior.Color = System.Drawing.Color.FromName(color);
            }
        }

        internal static void RestoreColor(Dictionary<Excel.Range, System.Drawing.Color> outliers)
        {
            foreach (KeyValuePair<Excel.Range, System.Drawing.Color> cell in outliers)
            {
                //Restore original color to cells flagged as outliers
                cell.Key.Interior.Color = cell.Value;
            }
        }

        //Error function (erf)
        public static double erf(double x)
        {
            //Save the sign of x
            int sign;
            if (x >= 0) sign = 1;
            else sign = -1;
            x = Math.Abs(x);
            //Constants
            double a1 = 0.254829592;
            double a2 = -0.284496736;
            double a3 = 1.421413741;
            double a4 = -1.453152027;
            double a5 = 1.061405429;
            double p = 0.3275911;

            double t = 1.0 / (1.0 + p * x);
            double y = 1.0 - (a1 * t + a2 * Math.Pow(t, 2.0) + a3 * Math.Pow(t, 3.0) + a4 * Math.Pow(t, 4.0) + a5 * Math.Pow(t, 5.0)) * Math.Exp(-x * x);
            //double y = 1.0 - (((((a5*t + a4)*t) + a3)*t + a2)*t + a1)*t*Math.Exp(-x*x);

            return sign * y;   //erf(-x) = -erf(x)
        }

        //Complementary error function (erfc)
        public static double erfc(double x)
        {
            return (1.0 - erf(x));
        }
    }
}
