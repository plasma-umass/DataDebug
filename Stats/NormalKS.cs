using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace DataDebug.Stats
{
    class NormalKS
    {
        private Excel.Range _cells;
        private int _size;
        private Double _mean;
        private Double _variance;
        private Double _standard_deviation;
        //private Dictionary<Excel.Range, Double> _error;
        //private List<Excel.Range> _ranked_errors;

        private Double __mean()
        {

            double sum = 0;
            foreach (Excel.Range cell in _cells)
            {
                sum += cell.Value;
            }
            return sum / _size;
        }

        private Double __standard_deviation()
        {
            return Math.Sqrt(_variance);
        }

        private Double __variance()
        {
            Double distance_sum_sq = 0;
            Double mymean = __mean();
            foreach (Excel.Range cell in _cells)
            {
                distance_sum_sq += Math.Pow(mymean - cell.Value, 2);
            }
            return distance_sum_sq / _size;
        }

        //Computes the phi function given a z-score. (This is the CDF for the normal distribution.)
        private static double __phi(double x)
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
                sign = -1;
            x = Math.Abs(x) / Math.Sqrt(2.0);

            // A&S formula 7.1.26
            double t = 1.0 / (1.0 + p * x);
            double y = 1.0 - (((((a5 * t + a4) * t) + a3) * t + a2) * t + a1) * t * Math.Exp(-x * x);

            return 0.5 * (1.0 + sign * y);
        }

        public NormalKS(Excel.Range r)
        {
            _cells = r;
            _size = r.Count;
            MessageBox.Show("Size: " + _size);
            _mean = __mean();
            _variance = __variance();
            _standard_deviation = __standard_deviation();
            MessageBox.Show("Standard deviation = " + _standard_deviation);
            MessageBox.Show("Mean = " + _mean);
            double[] cellsArray = new double[_size];
            int i = 0;
            foreach (Excel.Range ce in _cells)
            {
                cellsArray[i] = ce.Value;
                i++;
            }
            Array.Sort(cellsArray);
            double d_statistic = 0.0;
            foreach (double val in cellsArray)
            {
                int counter = 0; //Keeps track of how many cells have values less than or equal to the current cell
                //Loop through the cells to find how many hava values less than or equal to the current cell
                foreach (double val2 in cellsArray)
                {
                    if (val2 <= val)
                    {
                        counter++;
                    }
                }
                //double cdf_actual = 0.5 * (1 + Utilities.erf((val - _mean) / Math.Sqrt(2 * Math.Pow(_standard_deviation, 2.0))));
                double z_score = (val - _mean) / _standard_deviation;
                double cdf_observed = (double)counter / (double)_size;
                double current_d = Math.Abs(cdf_observed - __phi(z_score));
                //MessageBox.Show("Counter = " + counter + 
                //"\nObserved CDF = " + cdf_observed +
                //"\ncurrent D = " + current_d);
                if (d_statistic < current_d)
                {
                    d_statistic = current_d;
                    //MessageBox.Show("Observed CDF(" + z_score + ") = " + cdf_observed + "\nPhi(" + z_score + ") = " + __phi(z_score));
                }
            }
            MessageBox.Show("D statistic: " + d_statistic);

            //Now we test the D statistic to see if we reject H0

            //Array storing the critical values at the alpha = 0.05 confidence level for the KS test
            //The array is indexed by the sample size n (critical_values[0] corresponds to n=1, critical_values[1] corresponds to n=2, etc.)
            double[] critical_values = { 0.9500, 0.7764, 0.6360, 0.5652, 0.5094, //n=5
                                           0.4680,0.4361,0.4096,0.3875,0.3687, //n=10
                                           0.3524,0.3382,0.3255,0.3142,0.3040, //n=15
                                           0.2947,0.2863,0.2785,0.2714,0.2647, //n=20
                                           0.2586,0.2528,0.2475,0.2424,0.2377, //n=25
                                           0.2332,0.2290,0.2250,0.2212,0.2176, //n=30
                                           0.2141,0.2108,0.2077,0.2047,0.2018, //n=35
                                           0.1991, 0.1965,0.1939,0.1915,0.1891  //n=40
                                       };
            if (_size <= 40 && _size > 0)
            {
                if (d_statistic <= critical_values[_size - 1])
                {
                    MessageBox.Show("Your selection appears to be normally distributed. (alpha = 0.05)");
                }
                else
                {
                    MessageBox.Show("Your selection DOES NOT appear to be normally distributed. (alpha = 0.05)");
                }
            }
            else if (_size <= 0)
            {
                MessageBox.Show("Please select at least one cell.");
            }
            else
            {
                if (d_statistic <= 1.22 / Math.Sqrt(_size))
                {
                    MessageBox.Show("Your selection appears to be normally distributed. (alpha = 0.05)");
                }
                else
                {
                    MessageBox.Show("Your selection DOES NOT appear to be normally distributed. (alpha = 0.05)");
                }
            }


        }
    }
}
