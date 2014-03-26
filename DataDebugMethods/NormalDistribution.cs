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
        //TODO Add checks for non-numeric data in all methods of this class
        private readonly Excel.Range _cells;
        private readonly int _size;
        private readonly Double _mean;
        private readonly Double _variance;
        private readonly Double _standard_deviation;
        private Dictionary<Excel.Range, Double> _error;
        public List<Excel.Range> _ranked_errors;
        private int numeric_count = 0;
        // PRIVATE METHODS
        private Dictionary<Excel.Range, Double> __error()
        {
            Dictionary<Excel.Range, Double> d = new Dictionary<Excel.Range, Double>();
            foreach (Excel.Range cell in _cells)
            {
                if (cell.Value2 != null)
                {
                    try
                    {
                        var val = System.Convert.ToDouble(cell.Value2);
                        if (Math.Abs(_mean - val) / _standard_deviation > 2.0)
                        {
                            d.Add(cell, (double)Math.Abs(_mean - val) / _standard_deviation);
                        }
                    }
                    catch { }
                }
            }
            return d;
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
            //return sum / _size;
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

        // PUBLIC METHODS
        public int Length()
        {
            return numeric_count;
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
                try
                {
                    r1 = app.Union(r1, range_node.getCOMObject());
                }
                catch { }
            }
            _cells = r1;
            _size = r1.Count;
            _mean = __mean();
            _variance = __variance();
            _standard_deviation = __standard_deviation();
            _error = __error();
            _ranked_errors = __rank_errors();
        }

        public Dictionary<Excel.Range, System.Drawing.Color> PeirceOutliers()
        {
            List<Excel.Range> removals = new List<Excel.Range>();

            int doubtful_n = 1;
            bool remove_more = true;    // controls whether we continue to test
            int last_removals = 0;      // how many did we remove the last time around?

            while (remove_more)
            {
                List<Excel.Range> new_removals = new List<Excel.Range>();
                double r = R(doubtful_n, Length());
                if (r == -1) break;  // TODO: analytic derivation of R; in case R is undefined
                double max_deviation = StandardDeviation() * r;

                // find as many cells with error > max_deviation as possible
                for (int i = 0; i < Length(); i++)
                {
                    Excel.Range cell = _ranked_errors[i];
                    if (_error[cell] > max_deviation) new_removals.Add(cell);
                }

                // stop removing or continue?
                if ((new_removals.Count - last_removals) > 0)
                {
                    // adjust doubtful_n
                    doubtful_n += (new_removals.Count - last_removals);
                    last_removals = new_removals.Count;
                }
                else
                {
                    remove_more = false;
                }

                removals.AddRange(new_removals);
            }
            // add removals to removalsDict for coloring-undo
            Dictionary<Excel.Range, System.Drawing.Color> removalsDict = new Dictionary<Excel.Range, System.Drawing.Color>();
            foreach (Excel.Range c in removals)
            {
                System.Drawing.Color color = System.Drawing.ColorTranslator.FromOle((int)c.Interior.Color);
                try
                {
                    removalsDict.Add(c, color);
                }
                catch
                {
                }
            }

            return removalsDict;
        }

        public static Double R(int num_outliers, int data_sz)
        {
            // returns R value or -1 if undefined
            double[,] values = new double[,] {  {0.000, 0.000, 0.000, 0.000, 0.000, 0.000, 0.000, 0.000, 0.000},
                                                {0.000, 0.000, 0.000, 0.000, 0.000, 0.000, 0.000, 0.000, 0.000},
                                                {0.000, 0.000, 0.000, 0.000, 0.000, 0.000, 0.000, 0.000, 0.000},
                                                {1.196, 0.000, 0.000, 0.000, 0.000, 0.000, 0.000, 0.000, 0.000},
                                                {1.383, 1.078, 0.000, 0.000, 0.000, 0.000, 0.000, 0.000, 0.000},
                                                {1.509, 1.200, 0.000, 0.000, 0.000, 0.000, 0.000, 0.000, 0.000},
                                                {1.610, 1.299, 1.099, 0.000, 0.000, 0.000, 0.000, 0.000, 0.000},
                                                {1.693, 1.382, 1.187, 1.022, 0.000, 0.000, 0.000, 0.000, 0.000},
                                                {1.763, 1.453, 1.261, 1.109, 0.000, 0.000, 0.000, 0.000, 0.000},
                                                {1.824, 1.515, 1.324, 1.178, 1.045, 0.000, 0.000, 0.000, 0.000},
                                                {1.878, 1.570, 1.380, 1.237, 1.114, 0.000, 0.000, 0.000, 0.000},
                                                {1.925, 1.619, 1.430, 1.289, 1.172, 1.059, 0.000, 0.000, 0.000},
                                                {1.969, 1.663, 1.475, 1.336, 1.221, 1.118, 1.009, 0.000, 0.000},
                                                {2.007, 1.704, 1.516, 1.379, 1.266, 1.167, 1.070, 0.000, 0.000},
                                                {2.043, 1.741, 1.554, 1.417, 1.307, 1.210, 1.120, 1.026, 0.000},
                                                {2.076, 1.775, 1.589, 1.453, 1.344, 1.249, 1.164, 1.078, 0.000},
                                                {2.106, 1.807, 1.622, 1.486, 1.378, 1.285, 1.202, 1.122, 1.039},
                                                {2.134, 1.836, 1.652, 1.517, 1.409, 1.318, 1.237, 1.161, 1.084},
                                                {2.161, 1.864, 1.680, 1.546, 1.438, 1.348, 1.268, 1.195, 1.123},
                                                {2.185, 1.890, 1.707, 1.573, 1.466, 1.377, 1.298, 1.226, 1.158},
                                                {2.209, 1.914, 1.732, 1.599, 1.492, 1.404, 1.326, 1.255, 1.190},
                                                {2.230, 1.938, 1.756, 1.623, 1.517, 1.429, 1.352, 1.282, 1.218},
                                                {2.251, 1.960, 1.779, 1.646, 1.540, 1.452, 1.376, 1.308, 1.245},
                                                {2.271, 1.981, 1.800, 1.668, 1.563, 1.475, 1.399, 1.332, 1.270},
                                                {2.290, 2.000, 1.821, 1.689, 1.584, 1.497, 1.421, 1.354, 1.293},
                                                {2.307, 2.019, 1.840, 1.709, 1.604, 1.517, 1.442, 1.375, 1.315},
                                                {2.324, 2.037, 1.859, 1.728, 1.624, 1.537, 1.462, 1.396, 1.336},
                                                {2.341, 2.055, 1.877, 1.746, 1.642, 1.556, 1.481, 1.415, 1.356},
                                                {2.356, 2.071, 1.894, 1.764, 1.660, 1.574, 1.500, 1.434, 1.375},
                                                {2.371, 2.088, 1.911, 1.781, 1.677, 1.591, 1.517, 1.452, 1.393},
                                                {2.385, 2.103, 1.927, 1.797, 1.694, 1.608, 1.534, 1.469, 1.411},
                                                {2.399, 2.118, 1.942, 1.812, 1.710, 1.624, 1.550, 1.486, 1.428},
                                                {2.412, 2.132, 1.957, 1.828, 1.725, 1.640, 1.567, 1.502, 1.444},
                                                {2.425, 2.146, 1.971, 1.842, 1.740, 1.655, 1.582, 1.517, 1.459},
                                                {2.438, 2.159, 1.985, 1.856, 1.754, 1.669, 1.597, 1.532, 1.475},
                                                {2.450, 2.172, 1.998, 1.870, 1.768, 1.683, 1.611, 1.547, 1.489},
                                                {2.461, 2.184, 2.011, 1.883, 1.782, 1.697, 1.624, 1.561, 1.504},
                                                {2.472, 2.196, 2.024, 1.896, 1.795, 1.711, 1.638, 1.574, 1.517},
                                                {2.483, 2.208, 2.036, 1.909, 1.807, 1.723, 1.651, 1.587, 1.531},
                                                {2.494, 2.219, 2.047, 1.921, 1.820, 1.736, 1.664, 1.600, 1.544},
                                                {2.504, 2.230, 2.059, 1.932, 1.832, 1.748, 1.676, 1.613, 1.556},
                                                {2.514, 2.241, 2.070, 1.944, 1.843, 1.760, 1.688, 1.625, 1.568},
                                                {2.524, 2.251, 2.081, 1.955, 1.855, 1.771, 1.699, 1.636, 1.580},
                                                {2.533, 2.261, 2.092, 1.966, 1.866, 1.783, 1.711, 1.648, 1.592},
                                                {2.542, 2.271, 2.102, 1.976, 1.876, 1.794, 1.722, 1.659, 1.603},
                                                {2.551, 2.281, 2.112, 1.987, 1.887, 1.804, 1.733, 1.670, 1.614},
                                                {2.560, 2.290, 2.122, 1.997, 1.897, 1.815, 1.743, 1.681, 1.625},
                                                {2.568, 2.299, 2.131, 2.006, 1.907, 1.825, 1.754, 1.691, 1.636},
                                                {2.577, 2.308, 2.140, 2.016, 1.917, 1.835, 1.764, 1.701, 1.646},
                                                {2.585, 2.317, 2.149, 2.026, 1.927, 1.844, 1.773, 1.711, 1.656},
                                                {2.592, 2.326, 2.158, 2.035, 1.936, 1.854, 1.783, 1.721, 1.666},
                                                {2.600, 2.334, 2.167, 2.044, 1.945, 1.863, 1.792, 1.730, 1.675},
                                                {2.608, 2.342, 2.175, 2.052, 1.954, 1.872, 1.802, 1.740, 1.685},
                                                {2.615, 2.350, 2.184, 2.061, 1.963, 1.881, 1.811, 1.749, 1.694},
                                                {2.622, 2.358, 2.192, 2.069, 1.972, 1.890, 1.820, 1.758, 1.703},
                                                {2.629, 2.365, 2.200, 2.077, 1.980, 1.898, 1.828, 1.767, 1.711},
                                                {2.636, 2.373, 2.207, 2.085, 1.988, 1.907, 1.837, 1.775, 1.720},
                                                {2.643, 2.380, 2.215, 2.093, 1.996, 1.915, 1.845, 1.784, 1.729},
                                                {2.650, 2.387, 2.223, 2.101, 2.004, 1.923, 1.853, 1.792, 1.737},
                                                {2.656, 2.394, 2.230, 2.109, 2.012, 1.931, 1.861, 1.800, 1.745},
                                                {2.663, 2.401, 2.237, 2.116, 2.019, 1.939, 1.869, 1.808, 1.753}
                                            };

            if (data_sz > values.Length)
            {
                throw new Exception("Peirce's R is undefined for n = " + data_sz + " and num_outliers = " + num_outliers);
            }
            double lookup = values[data_sz, num_outliers - 1]; // need to adjust because outliers start count from 0 in table

            if (lookup == 0.000)
            {
                return -1;
            }
            return lookup;
        }

        public Double StandardDeviation()
        {
            return _standard_deviation;
        }

        public Double Variance()
        {
            return _variance;
        }

        public Excel.Range WorstError()
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

        public Excel.Range getError(int rank)
        {
            return _ranked_errors[rank];
        }

        public int errorsCount()
        {
            return _ranked_errors.Count;
        }
    }
}
