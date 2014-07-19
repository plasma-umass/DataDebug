using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using TreeNode = DataDebugMethods.TreeNode;
using TreeScore = System.Collections.Generic.Dictionary<DataDebugMethods.TreeNode, int>;


namespace UserSimulation
{
    public abstract class CutoffKind {
        public abstract List<KeyValuePair<TreeNode,int>> applyCutoff(TreeScore scores, HashSet<AST.Address> known_good);
        public abstract double Threshold
        {
            get;
        }
        public virtual bool isCountBased
        {
            get { return false; }
        }
    }

    public class NormalCutoff : CutoffKind
    {
        double _t;
        public NormalCutoff(double pct_threshold)
        {
            _t = pct_threshold;
        }
        override public double Threshold
        {
            get { return _t; }
        }
        override public List<KeyValuePair<TreeNode, int>> applyCutoff(TreeScore scores, HashSet<AST.Address> known_good)
        {
            //Using an outlier test for highlighting 
            //scores that fall outside of two standard deviations from the others
            //The one-sided 5% cutoff for the normal distribution is 1.6448.

            var scores_list = scores.OrderByDescending(pair => pair.Value).ToList();

            List<KeyValuePair<TreeNode, int>> filtered_high_scores = null;

            //Code for doing normal outlier analysis on the scores:
            //find mean:
            double sum = 0.0;
            foreach (double d in scores.Values)
            {
                sum += d;
            }
            double mean = sum / scores.Values.Count;
            //find variance
            double distance_sum_sq = 0.0;
            foreach (double d in scores.Values)
            {
                distance_sum_sq += Math.Pow(mean - d, 2);
            }
            double variance = distance_sum_sq / scores.Values.Count;

            //find std. deviation
            double std_deviation = Math.Sqrt(variance);

            if (_t == 0.05)
            {
                filtered_high_scores = scores_list.Where(kvp => kvp.Value > mean + std_deviation * 1.6448).ToList();
            }
            else if (_t == 0.1)   //10% cutoff 1.2815
            {
                filtered_high_scores = scores_list.Where(kvp => kvp.Value > mean + std_deviation * 1.2815).ToList();
            }
            else if (_t == 0.025) //2.5% cutoff 1.9599
            {
                filtered_high_scores = scores_list.Where(kvp => kvp.Value > mean + std_deviation * 1.9599).ToList();
            }
            else if (_t == 0.075) //7.5% cutoff 1.4395
            {
                filtered_high_scores = scores_list.Where(kvp => kvp.Value > mean + std_deviation * 1.4395).ToList();
            }
            else
            {
                throw new Exception("Uhhh.... What's my cutoff?");
            }

            return filtered_high_scores;
        }
    }

    public class QuantileCutoff : CutoffKind
    {
        double _t;
        public QuantileCutoff(double pct_threshold)
        {
            _t = pct_threshold;
        }
        override public double Threshold
        {
            get { return _t; }
        }
        override public List<KeyValuePair<TreeNode, int>> applyCutoff(TreeScore scores, HashSet<AST.Address> known_good)
        {
            var scores_list = scores.OrderByDescending(pair => pair.Value).ToList();

            int start_ptr = 0;
            int end_ptr = 0;

            List<KeyValuePair<TreeNode, int>> high_scores = new List<KeyValuePair<TreeNode, int>>();

            while ((double)start_ptr / scores_list.Count < _t) //the start of this score region is before the cutoff
            {
                // make sure that we don't go off the end of the list
                if (start_ptr >= scores_list.Count)
                {
                    break;
                }

                //while the scores at the start and end pointers are the same, bump the end pointer
                while (end_ptr < scores_list.Count && scores_list[start_ptr].Value == scores_list[end_ptr].Value)
                {
                    end_ptr++;
                }
                //Now the end_pointer points to the first index with a lower score
                //If the number of entries with the current value is fewer than the significance cutoff, add all values of this score to the high_scores list; the number of entries is equal to the end_ptr since end_ptr is zero-based
                //There is some added "wiggle room" to the cutoff, so that the last entry is allowed to straddle the cutoff bound.
                //  To do this, we add (1 / total number of entries) to the cutoff
                //The purpose of the wiggle room is to allow us to deal with small ranges (less than 20 entries), since a single entry accounts
                //for more than 5% of the total.
                if ((double)end_ptr / scores_list.Count < _t + (double)1.0 / scores_list.Count)
                {
                    //add all values of the current score to high_scores list
                    for (; start_ptr < end_ptr; start_ptr++)
                    {
                        high_scores.Add(scores_list[start_ptr]);
                    }
                    //Increment the start pointer to the start of the next score region
                    start_ptr++;
                }
                else    //if this score region extends past the cutoff, we don't add any of its values to the high_scores list, and stop
                {
                    break;
                }
            }
            // filter out cells marked as OK
            return high_scores.Where(kvp => !known_good.Contains(kvp.Key.GetAddress())).ToList();
        }
    }

    public class NumberCutoff : CutoffKind
    {
        int _n;
        public NumberCutoff(int count)
        {
            _n = count;
        }
        override public double Threshold
        {
            get { return _n; }
        }
        override public bool isCountBased
        {
            get { return true; }
        }
        override public List<KeyValuePair<TreeNode, int>> applyCutoff(TreeScore scores, HashSet<AST.Address> known_good)
        {
            // sort scores, biggest first
            var scores_list = scores.OrderByDescending(pair => pair.Value).ToList();

            // return the n largest (and force eager evaluation)
            // excluding cells marked as OK
            return scores_list.Take(_n).Where(kvp => !known_good.Contains(kvp.Key.GetAddress())).ToList();
        }
    }
}
