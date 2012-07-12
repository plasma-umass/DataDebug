using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;

namespace DataDebug.Stats
{
    class NormalAD
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

        //Standard deviation of the sample
        private Double __standard_deviation()
        {
            return Math.Sqrt(_variance);
        }

        //Variance of the sample
        private Double __variance()
        {
            Double distance_sum_sq = 0;
            Double mymean = __mean();
            foreach (Excel.Range cell in _cells)
            {
                distance_sum_sq += Math.Pow(mymean - cell.Value, 2);
            }
            return distance_sum_sq / (_size - 1);
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

        public NormalAD(Excel.Range r)
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
            double ad_statistic = 0.0;
            for (int j = 0; j < cellsArray.Length; j++)
            {
                double val = cellsArray[j];
                double z_score = (val - _mean) / _standard_deviation;
                double z_score2 = (cellsArray[cellsArray.Length - j - 1] - _mean) / _standard_deviation;
                //AD = SUM[i=1 to n] (1-2i)/n * {ln(F0[z_i]) + ln(1-F0[Z_(n+1-i)]) } - n
                //Reject if AD > CV = 0.752 / (1 + 0.75/n + 2.25/(n^2) )
                double s = (1 - 2 * (j + 1)) * (Math.Log(__phi(z_score)) + Math.Log(1 - __phi(z_score2)));
                ad_statistic = ad_statistic + s;
                //MessageBox.Show("s: " + s) ; //LN(F[z]) = " + Math.Log(__phi(z_score)) + "; LN(F[z2]) = " + Math.Log(1 - __phi(z_score2)));
                //MessageBox.Show(j + 1 + ": " + s); 

            }
            //MessageBox.Show("Sum: " + (ad_statistic));
            ad_statistic = ad_statistic / cellsArray.Length;
            ad_statistic = ad_statistic - cellsArray.Length;
            MessageBox.Show("A-D statistic: " + ad_statistic);
            //MessageBox.Show("Adjusted A-D statistic: " + ad_statistic * (1 + 0.75/cellsArray.Length + 2.25



            //Now we test the D statistic to see if we reject H0
            double critical_value = 0.752 / (1 + 0.75 / cellsArray.Length + 2.25 / (cellsArray.Length * cellsArray.Length));
            MessageBox.Show("CV = " + critical_value);
            if (_size > 0)
            {
                if (ad_statistic <= critical_value)
                {
                    MessageBox.Show("Your selection does not show significant deviation from a normal distribution. (alpha = 0.05)");
                }
                else
                {
                    MessageBox.Show("Your selection DOES NOT appear to be normally distributed. (alpha = 0.05)");
                }
            }
            else //if (_size <= 0)
            {
                MessageBox.Show("Please select at least one cell.");
            }

        }
    }
}
