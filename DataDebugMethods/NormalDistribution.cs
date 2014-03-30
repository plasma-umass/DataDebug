using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;


namespace DataDebugMethods
{
    public class NormalDistribution
    {
        private readonly Excel.Range _cells;
        private readonly int _size;
        private readonly Double _mean;
        private readonly Double _variance;
        private readonly Double _standard_deviation;
        private Dictionary<Excel.Range, Double> _error;
        public List<Excel.Range> _ranked_errors;
        private int numeric_count = 0;

        private Dictionary<Excel.Range, Double> __error()
        {
            Dictionary<Excel.Range, Double> error_dict = new Dictionary<Excel.Range, Double>();
            foreach (Excel.Range cell in _cells)
            {
                if (cell.Value2 != null)
                {
                    try
                    {
                        var val = System.Convert.ToDouble(cell.Value2);
                        if (Math.Abs(_mean - val) / _standard_deviation > 2.0)
                        {
                            error_dict.Add(cell, (double)Math.Abs(_mean - val) / _standard_deviation);
                        }
                    }
                    catch { }
                }
            }
            return error_dict;
        }

        private Double __mean()
        {
            double sum = 0.0;
            foreach (Excel.Range cell in _cells)
            {
                if (cell.Value2 != null)
                {
                    try
                    {
                        var val = System.Convert.ToDouble(cell.Value2);
                        sum += val;
                        numeric_count++;
                    } catch { }
                }
            }
            return sum / numeric_count;
        }

        private List<Excel.Range> __rank_errors()
        {
            List<Excel.Range> rs = _error.OrderBy(pair => pair.Value).Select(pair => pair.Key).ToList<Excel.Range>();
            rs.Reverse();
            return rs;
        }

        private Double __standard_deviation()
        {
            return Math.Sqrt(_variance);
        }

        private Double __variance()
        {
            Double distance_sum_sq = 0;
            Double mymean = Mean();
            foreach (Excel.Range cell in _cells)
            {
                if (cell.Value2 != null)
                {
                    try
                    {
                        var val = System.Convert.ToDouble(cell.Value2);
                        distance_sum_sq += Math.Pow(mymean - val, 2);
                    } catch { }
                }
            } 
            return distance_sum_sq / numeric_count;
        }

        public Double Mean()
        {
            return _mean;
        }

        public NormalDistribution(Excel.Range r)
        {
            _cells = r;
            _size = r.Count;
            _mean = __mean();
            _variance = __variance();
            _standard_deviation = __standard_deviation();
            _error = __error();
            _ranked_errors = __rank_errors();
        }

        public NormalDistribution(TreeNode[] range_nodes, Excel.Application app)
        {
            //turn the dictionary into an Excel.Range
            Excel.Range r1 = range_nodes.First().getCOMObject(); 
            foreach (TreeNode range_node in range_nodes)
            {
                try  // in a try-catch because Union malfunctioned in one case
                {
                    r1 = app.Union(r1, range_node.getCOMObject());
                } catch { }
            }
            _cells = r1;
            _size = r1.Count;
            _mean = __mean();
            _variance = __variance();
            _standard_deviation = __standard_deviation();
            _error = __error();
            _ranked_errors = __rank_errors();
        }

        public Double getStandardDeviation()
        {
            return _standard_deviation;
        }

        public Double getVariance()
        {
            return _variance;
        }

        public Excel.Range getWorstError()
        {
            return _ranked_errors.First();
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

        //Computes the phi function given a z-score. (This is the CDF for the normal distribution.)
        public static double __phi(double x)
        {
            // constants
            double a1 = 0.254829592;
            double a2 = -0.284496736;
            double a3 = 1.421413741;
            double a4 = -1.453152027;
            double a5 = 1.061405429;
            double p = 0.3275911;

            // Save the sign of x
            int sign = 1;
            if (x < 0)
            {
                sign = -1;
            }
            x = Math.Abs(x) / Math.Sqrt(2.0);

            // A&S formula 7.1.26
            double t = 1.0 / (1.0 + p * x);
            double y = 1.0 - (((((a5 * t + a4) * t) + a3) * t + a2) * t + a1) * t * Math.Exp(-x * x);
 
            return 0.5 * (1.0 + sign * y);
         }

        public List<Excel.Range> getRankedErrors()
        {
            return _ranked_errors;
        }

        public Excel.Range getErrorAtPosition(int rank)
        {
            return _ranked_errors[rank];
        }

        public int getErrorsCount()
        {
            return _ranked_errors.Count;
        }
    }
}
