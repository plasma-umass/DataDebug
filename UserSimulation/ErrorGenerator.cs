using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.FSharp.Core;
using Sign = LongestCommonSubsequence.Sign;
using LCSError = LongestCommonSubsequence.Error;
using ErrorString = System.Tuple<string, System.Collections.Generic.List<LongestCommonSubsequence.Error>>;
using OptChar = Microsoft.FSharp.Core.FSharpOption<char>;
using CellDict = System.Collections.Generic.Dictionary<AST.Address, string>;

namespace UserSimulation
{
    public class ErrorGenerator
    {
        private Random r = new Random();
        //Keeps the distributions that have been generated so far, so that they don't have to be generated again later
        private Dictionary<OptChar, Dictionary<string,double>> _char_distributions_dict = new Dictionary<OptChar,Dictionary<string,double>>();

        private Dictionary<int, double> _transpositions_distribution_dict = new Dictionary<int, double>();

        //Gets the distribution of strings for a particular character
        //DOES NOT use previously generated distributions; generates the distribution every time
        private Dictionary<string, double> GetDistributionOfStringsForChar(OptChar c, Classification classification)
        {
            OptChar key = c;
            Dictionary<string, double> distribution;
            //Generate the probability distribution based on the classification, which contains counts of observations
            distribution = GenerateDistributionForChar(key, classification);
            //If our dictionary does not have any information about this character, we return the character with probability 1.0
            if (distribution.Count == 0)
            {
                distribution.Add("" + c.Value, 1.0);
            }
            return distribution;
        }


        //Gets the distribution of strings for a particular character
        //If the distribution has been generated before, it is reused from the _char_distributions_dict
        private Dictionary<string, double> GetDistributionOfStringsForCharReuse(OptChar c, Classification classification)
        {
            OptChar key = c;
            Dictionary<string, double> distribution;
            //if we have already generated a distribution for this character, return it
            if (_char_distributions_dict.TryGetValue(key, out distribution))
            {
                return distribution;
            }
            //otherwise generate the distribution and then return it
            else
            {
                distribution = GenerateDistributionForChar(key, classification);
                //If our dictionary does not have any information about this character, we return the character with probability 1.0
                if (distribution.Count == 0)
                {
                    distribution.Add("" + c.Value, 1.0);
                }
                _char_distributions_dict.Add(key, distribution);
                return distribution;
            }
        }

        public string[] GenerateErrorStrings(string orig, Classification c, int k)
        {
            var e = Enumerable.Range(0, k);
            var strs = e.AsParallel().Select( i => GenerateErrorString(orig, c) );
            return strs.ToArray();
        }

        public CellDict RandomlyGenerateErrors(CellDict original_inputs, Classification c, double proportion)
        {
            // number of cells to typo
            int n = (int)Math.Ceiling(original_inputs.Count * proportion);

            var introduced_errors = new CellDict();
            var oi = original_inputs.ToArray();

            int i = 0;
            while (i < n)
            {
                int idx = r.Next(oi.Length);
                KeyValuePair<AST.Address,String> cell = oi[idx];
                if (introduced_errors.ContainsKey(cell.Key) || cell.Value == null)
                {   // if we've already typo'ed this value, or if the value of the cell is null move on
                    continue;
                }
                else
                {
                    //For verifying the error distribution produced by the generator
                    //Dictionary<string, int> errors_dict = new Dictionary<string, int>();
                    //for (int x = 0; x < 2000; x++)
                    //{
                    //    string error = GenerateErrorString(cell.Value, c);
                    //    if (errors_dict.ContainsKey(error))
                    //    {
                    //        //update the value
                    //        int count;
                    //        if (errors_dict.TryGetValue(error, out count))
                    //        {
                    //            errors_dict.Remove(error);
                    //            errors_dict.Add(error, count + 1);
                    //        }
                    //    }
                    //    else
                    //    {
                    //        //add the string with value 1
                    //        errors_dict.Add(error, 1);
                    //    }
                    //}
                    introduced_errors.Add(cell.Key, GenerateErrorString(cell.Value, c));
                }
                i++;
            }

            return introduced_errors;
        }

        //Generates the distribution of strings for a particular character given a classification
        private Dictionary<string, double> GenerateDistributionForChar(OptChar c, Classification classification)
        {
            var typo_dict = classification.GetTypoDict();
            var kvps = typo_dict.Where(pair => {
                if (OptChar.get_IsNone(pair.Key.Item1))
                {
                    if (OptChar.get_IsNone(c))
                    {
                        return true;
                    }
                    return false;
                }
                else
                {
                    return pair.Key.Item1.Equals(c);
                }
            }).ToArray();
            var sum = kvps.Select(pair => pair.Value).Sum();
            var distribution = kvps.Select(pair => new KeyValuePair<string,double>(pair.Key.Item2, (double) pair.Value / sum));
            return distribution.ToDictionary(pair => pair.Key, pair => pair.Value);
        }

        /// <summary>
        /// Given a distribution, this method chooses a string from the distribution at random
        /// based on the probabilities given in the distribution.
        /// </summary>
        private string GetRandomStringFromDistribution(Dictionary<string, double> distribution)
        {
            var rng = new Random();
            var rand = rng.NextDouble();

            int i = 0;
            double sum = distribution.ElementAt(i).Value;
            while (sum < rand)
            {
                i++;
                sum += distribution.ElementAt(i).Value;
            }

            var kvp = distribution.ElementAt(i);
            return kvp.Key;
        }

        private Dictionary<int, double> GetDistributionOfTranspositions(Classification classification)
        {
            //if we have already generated a distribution, return it
            if (_transpositions_distribution_dict.Count != 0)
            {
                return _transpositions_distribution_dict;
            }
            else //otherwise generate the distribution and then return it
            {
                _transpositions_distribution_dict = GenerateTranspositionsDistribution(classification);
                //If our dictionary does not have any information about transpositions, we add to it delta = 0 with probability 1.0
                if (_transpositions_distribution_dict.Count == 0)
                {
                    _transpositions_distribution_dict.Add(0, 1.0);
                }
                return _transpositions_distribution_dict;
            }
        }

        private Dictionary<int, double> GenerateTranspositionsDistribution(Classification classification)
        {
            var transposition_dict = classification.GetTranspositionDict();
            var sum = transposition_dict.Select(pair => pair.Value).Sum();
            var distribution = transposition_dict.Select(pair => new KeyValuePair<int, double>(pair.Key, (double)pair.Value / sum));
            return distribution.ToDictionary(pair => pair.Key, pair => pair.Value);
        }
        
        private int GetRandomTranspositionFromDistribution(Dictionary<int, double> transposition_distribution)
        {
            var rng = new Random();
            var rand = rng.NextDouble();

            int i = 0;
            double sum = transposition_distribution.ElementAt(i).Value;
            while (sum < rand)
            {
                i++;
                sum += transposition_distribution.ElementAt(i).Value;
            }

            var kvp = transposition_distribution.ElementAt(i);
            return kvp.Key;
        }

        private static string OptCharToString(OptChar ch)
        {
            if (OptChar.get_IsSome(ch))
            {
                return ch.Value.ToString();
            }
            else
            {
                return "";
            }
        }

        /// <summary>
        /// Samples randomly from a multinomial probability vector.
        /// </summary>
        /// <param name="probabilities">A double[] containing p values; must sum to 1!</param>
        /// <param name="r">A random number generator</param>
        /// <returns></returns>
        public int MultinomialSample(double[] probabilities)
        {
            const double EPSILON = 0.01;

            // normalize probabilities
            var k = 1.0 / probabilities.Sum();
            var probs = probabilities.Select(pr => k * pr);

            if (probs.Sum() < 1 - EPSILON || probs.Sum() > 1 + EPSILON)
            {
                throw new Exception("Probabilities do not sum to 1.");
            }

            // draw intervals
            double[] intervals = probs.Select((pr_1, i) => probs.Where((pr_2, j) => j < i).Sum() + pr_1).ToArray();

            // draw a sample
            var s = r.NextDouble();

            // the inputchars at index idx is the char to mistype
            for (int i = 0; i < intervals.Length; i++)
            {
                if (s > intervals[i])
                {
                    continue;
                }
                else
                {
                    return i;
                }
            }

            throw new Exception("Cannot find appropriate bin.");
        }

        /// <summary>
        /// Transpose characters.  This function expects an OptChar array with no
        /// leading or trailing empty characters, i.e., OptChar.None.  If guar == -1
        /// then there are no guaranteed transpositions, otherwise guar is the index
        /// of the element that must be transposed.
        /// </summary>
        /// <param name="input"></param>
        /// <param name="transpositions"></param>
        /// <param name="guar"></param>
        /// <returns>The input array with transpositions applied.</returns>
        public OptChar[] Transposize(OptChar[] input,
                                     Dictionary<int,int> transpositions,
                                     int guar)
        {
            // copy input to new array
            OptChar[] output = new OptChar[input.Length];
            Array.Copy(input, output, input.Length);

            // for each character in the string, sample from the transposition dict
            for (int i = 0; i < input.Length; i++)
            {
                // condition on the number of possible transpositions
                // to the right and to the left
                // if the character is a "guaranteed transposition", ensure that
                // the 0-transposition is not in the distribution
                var dist = transpositions.Where(kvp => kvp.Key < input.Length - i
                                                       && kvp.Key >= -i
                                                       && i == guar ? kvp.Key != 0 : true)
                                         .ToDictionary(pair => pair.Key, pair => pair.Value);
                var counts = Enumerable.Range(-i, input.Length)
                                       .Select(z => dist.ContainsKey(z) ? dist[z] : 0);
                var total = counts.Sum();
                if (total != 0)
                {
                    var prs = counts.Select(z => (double)z / total).ToArray();
                    // sample (in this case, bins always start at zero and are in order,
                    // so j == # of transpositions to right)
                    var j = MultinomialSample(prs);
                    // swap chars
                    OptChar ith = output[i];
                    output[i] = output[j];
                    output[j] = ith;
                }
            }

            return output;
        }

        public string Typoize(OptChar[] input,
                              Dictionary<Tuple<OptChar, string>, int> typos,
                              int guar)
        {
            var output = "";

            // for each character in the string, sample from the typo dict
            for (int i = 0; i < input.Length; i++)
            {
                KeyValuePair<Tuple<OptChar, String>, int>[] dist;
                // handle case where the input character is an empty char
                // and condition on the possible typos for this optchar
                if (OptChar.get_IsNone(input[i]))
                {
                    // the input character is the empty char, so condition on empty chars
                    var dist_1 = typos.Where(kvp => kvp.Key.Item1 == null);
                    // if the current character is a guaranteed typo, ensure that
                    // an empty character does not appear in the output
                    dist = (i == guar ? dist_1.Where(kvp => !kvp.Key.Item2.Equals("")) : dist_1).ToArray();
                }
                else
                {
                    // condition on the possible typos for this particular OptChar
                    var dist_1 = typos.Where(kvp => input[i].Equals(kvp.Key.Item1));
                    // get the string corresponding to the current OptChar
                    var str_i = OptChar.get_IsNone(input[i]) ? "" : input[i].Value.ToString();

                    // if the current character is a guaranteed typo, ensure that
                    // the conditioned OptChar does not appear in the output
                    dist = (i == guar ? dist_1.Where(kvp => !kvp.Key.Item2.Equals(str_i)) : dist_1).ToArray();
                }
                
                var total = dist.Select(kvp => kvp.Value).Sum();
                var prs = dist.Select(kvp => (double)kvp.Value / total).ToArray();
                if (prs.Length == 0)
                {
                    output += OptCharToString(input[i]);
                }
                else
                {
                    // sample
                    var j = MultinomialSample(prs);
                    // j corresponds to what typo string?
                    output += dist[j].Key.Item2;
                }
            }

            return output;
        }

        public OptChar[] StringToOptCharArray(string input)
        {
            if (input.Length == 0)
            {
                return new [] { OptChar.None };
            }
            else
            {
                return input.ToCharArray().Select(ch => new OptChar(ch)).ToArray();
            }
        }

        public OptChar[] AddLeadingTrailingSpace(OptChar[] input)
        {
            List<OptChar> trailing = input.ToList();
            trailing.Add(OptChar.None);                                 // add trailing empty string
            List<OptChar> leading = new[] { OptChar.None }.ToList();    // add leading empty string
            return leading.Concat(trailing).ToArray();
        }

        public string GenerateErrorString(string input, Classification c)
        {
            // get typo dict
            var td = c.GetTypoDict();

            // get transposition dict
            var trd = c.GetTranspositionDict();

            // convert the input into a char array
            var ochars = StringToOptCharArray(input);

            // add leading and trailing 'empty characters'
            var inputchars = AddLeadingTrailingSpace(ochars);

            // calculate the marginal probabilities of NOT making a typo for each char in input
            double[] PrsCharNotTypo = inputchars.Select(oc =>
            {
                var key = new Tuple<OptChar, string>(oc, OptCharToString(oc));
                int count;
                if (!td.TryGetValue(key, out count)) {
                    count = 0;
                }
                // funny case to handle the fact that FSharpOption.None == null
                var cond_dist = td.Where(kvp => kvp.Key.Item1 == null ? oc == null : kvp.Key.Item1.Equals(oc));
                int total = cond_dist.Aggregate(0, (acc, kvp) => acc + kvp.Value);
                if (total == 0)
                {
                    return 1.0;
                }
                else
                {
                    return (double)count / total;
                }
            }).ToArray();

            // calculate the probability of making at least one error
            // might need log-probs here
            double PrTypo = 1.0 - PrsCharNotTypo.Aggregate(1.0, (acc, pr_not_typo) => acc * pr_not_typo);

            // calculate the marginal probabilities of NOT making a
            // transposition for each position in the input
            // note that we do NOT consider the empty strings here
            // For strings of length 1, the probability of not making a
            // transposition should be exactly 1. 
            double[] PrsPosNotTrans = ochars.Length > 1 ? ochars.ToArray().Select((oc, idx) =>
            {
                int count;
                if (!trd.TryGetValue(0, out count)) {
                    count = 0;
                }
                int total = trd.Where(kvp => kvp.Key < input.Length - idx && kvp.Key >= -idx).Select(kvp => kvp.Value).Sum();
                if (total == 0)
                {
                    return 1.0;
                }
                else
                {
                    return (double)count / total;
                }
            }).ToArray() : new [] { 1.0 };

            // calculate the probability of having at least one transposition
            double PrTrans = 1.0 - PrsPosNotTrans.Aggregate(1.0, (acc, pr_not_trans) => acc * pr_not_trans);

            // calculate the relative probability of making a typo vs a transposition
            double RelPrTypo = PrTypo / (PrTypo + PrTrans);

            // init with original input in case typos/transpositions prove to be impossible
            string output = input;

            // the while loop ensures that we do not return an unmodified string.
            // for most strings, returning an unmodified string is very unlikely
            do
            {
                // flip a coin to determine whether our guaranteed error is a typo or a transposition
                if (r.NextDouble() < RelPrTypo)
                {   // is a typo
                    // determine the index of the guaranteed typo
                    double[] PrsMistype = PrsCharNotTypo.Select(pr => 1.0 - pr).ToArray();
                    // if there are no possible typos then we just can't produce one
                    if (PrsMistype.Sum() == 0)
                    {
                        break;
                    }
                    var i = MultinomialSample(PrsMistype);
                    // run transposition algorithm & add leading/trailing empty chars
                    // we set the guaranteed transposition index to -1 to ensure that no
                    // transpositions are guaranteed
                    OptChar[] input_t = AddLeadingTrailingSpace(Transposize(ochars, trd, -1));
                    // run typo algorithm (adjust i for leading space)
                    output = Typoize(input_t, td, i);
                }
                else
                {   // is a transposition
                    // determine the index of the guaranteed transposition
                    double[] PrsMistype = PrsPosNotTrans.Select(pr => 1.0 - pr).ToArray();
                    // if there are no possible transpositions then we just can't produce one
                    if (PrsMistype.Sum() == 0)
                    {
                        break;
                    }
                    var i = MultinomialSample(PrsMistype);
                    // run transposition algorithm & add leading/trailing empty chars
                    OptChar[] input_t = AddLeadingTrailingSpace(Transposize(ochars, trd, i));
                    // run typo algorithm; set guaranteed typo index to -1 to ensure that no
                    // typo is guaranteed
                    output = Typoize(input_t, td, -1);
                }
            } while (input == output);

            return output;
        }
    }
}
