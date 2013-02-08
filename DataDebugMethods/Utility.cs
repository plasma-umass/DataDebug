using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
using System.Text.RegularExpressions;

namespace DataDebugMethods
{
    public static class Utility
    {
        public static AddressRange AddressOfXLRange(Excel.Range rng)
        {
            
            var rng_r1c1 = rng.Address[true, true, Excel.XlReferenceStyle.xlR1C1, false];
            if (rng_r1c1.Split(':').Count() == 2)
            {
                var address_matcher = new Regex("^R([0-9]+)C([0-9]+):R([0-9]+)C([0-9]+)$", RegexOptions.Compiled);
                var mo = address_matcher.Match(rng_r1c1);
                return new AddressRange(System.Convert.ToInt32(mo.Groups[1].Value),
                                        System.Convert.ToInt32(mo.Groups[2].Value),
                                        System.Convert.ToInt32(mo.Groups[3].Value),
                                        System.Convert.ToInt32(mo.Groups[4].Value)
                                       );
            }
            else
            {
                var address_matcher = new Regex("^R([0-9]+)C([0-9]+)$", RegexOptions.Compiled);
                var mo = address_matcher.Match(rng_r1c1);
                return new AddressRange(System.Convert.ToInt32(mo.Groups[1].Value),
                                        System.Convert.ToInt32(mo.Groups[2].Value),
                                        System.Convert.ToInt32(mo.Groups[1].Value),
                                        System.Convert.ToInt32(mo.Groups[2].Value)
                                       );
            }
        }

        public static bool InsideRectangle(Excel.Range rng, AddressRange rect)
        {
            var addr_rng = AddressOfXLRange(rng);

            var is_bad = (addr_rng.getXLeft() < rect.getXLeft() ||
                          addr_rng.getYTop() < rect.getYTop() ||
                          addr_rng.getXRight() > rect.getXRight() ||
                          addr_rng.getYBottom() > rect.getYBottom());

            return !is_bad;
        }

        public static bool InsideUsedRange(Excel.Range rng)
        {
            return InsideRectangle(rng, UsedRange(rng));
        }

        public static AddressRange UsedRange(Excel.Range rng)
        {
            return AddressOfXLRange(rng.Worksheet.UsedRange);
        }
    }

    public class AddressRange
    {
        private int _x_left;
        private int _y_top;
        private int _x_right;
        private int _y_bottom;

        public AddressRange(int x_left, int y_top, int x_right, int y_bottom)
        {
            _x_left = x_left;
            _y_top = y_top;
            _x_right = x_right;
            _y_bottom = y_bottom;
        }

        public int getXLeft() { return _x_left; }
        public int getYTop() { return _y_top; }
        public int getXRight() { return _x_right; }
        public int getYBottom() { return _y_bottom; }
        public Tuple<int, int> getTopLeft() { return new Tuple<int,int>(_x_left, _y_top); }
        public Tuple<int, int> getBottomRight() { return new Tuple<int, int>(_x_right, _y_bottom); }
        public string ToString()
        {
            return "(" + _x_left.ToString() + "," + _y_top.ToString() + "),(" + _x_right.ToString() + "," + _y_bottom.ToString() + ")";
        }
    }
}
