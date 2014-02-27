using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.FSharp.Core;
using Sign = LongestCommonSubsequence.Sign;
using LCSError = LongestCommonSubsequence.Error;
using ErrorString = System.Tuple<string, System.Collections.Generic.List<LongestCommonSubsequence.Error>>;
using OptChar = Microsoft.FSharp.Core.FSharpOption<char>;

namespace UserSimulation
{
    public class ErrorGenerator
    {
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
            var strs = e.AsParallel().Select( i => {
                var outstr = "";
                while ((outstr = GenerateErrorString(orig, c).Item1).Equals(orig))
                {

                }
                return outstr;
            });
            return strs.ToArray();
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
        public int MultinomialSample(double[] probabilities, Random r)
        {
            const double EPSILON = 0.01;
            System.Diagnostics.Debug.Assert(probabilities.Sum() > 1 - EPSILON && probabilities.Sum() < 1 + EPSILON);

            // draw intervals
            double[] intervals = probabilities.Select((pr_1, i) => probabilities.Where((pr_2, j) => j < i).Sum() + pr_1).ToArray();

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

        public OptChar[] Transposize(OptChar[] input,
                                     Dictionary<int,int> transpositions,
                                     int guar,
                                     Random r)
        {
            // strip leading and trailing whitspace
            OptChar[] output = new OptChar[input.Length - 2];
            Array.Copy(input, 1, output, 0, input.Length - 1);

            // for each character in the string, sample from the transposition dict
            for (int i = 0; i < input.Length; i++)
            {
                // condition on the number of possible transpositions to the right
                // if the character is a "guaranteed transposition", ensure that
                // the 0-transposition is not in the distribution
                var dist = transpositions.Where(kvp => kvp.Key < input.Length - i
                                                       && i == guar ? kvp.Key != 0 : true);
                var total = dist.Select(kvp => kvp.Value).Sum();
                var prs = dist.Select(kvp => (double)kvp.Value / total).ToArray();
                // sample (in this case, bins always start at zero and are in order,
                // so j == # of transpositions to right)
                var j = MultinomialSample(prs, r);
                // swap chars
                OptChar ith = output[i];
                output[i] = output[j];
                output[j] = ith;
            }

            return output;
        }

        public string Typoize(OptChar[] input,
                              Dictionary<Tuple<OptChar, string>, int> typos,
                              int guar,
                              Random r)
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
                    // if the current character is a guaranteed typo, ensure that
                    // the conditioned OptChar does not appear in the output
                    dist = (i == guar ? dist_1.Where(kvp => !kvp.Key.Item2.Equals(input[i])) : dist_1).ToArray();
                }
                
                var total = dist.Select(kvp => kvp.Value).Sum();
                var prs = dist.Select(kvp => (double)kvp.Value / total).ToArray();
                // sample
                var j = MultinomialSample(prs, r);
                // j corresponds to what typo string?
                output += dist[j].Key.Item2;
            }

            return output;
        }

        public OptChar[] StringToOptCharArray(string input)
        {
            return input.ToCharArray().Select(ch => new OptChar(ch)).ToArray();
        }

        public OptChar[] AddLeadingTrailingSpace(OptChar[] input)
        {
            input.ToList().Add(OptChar.None);                           // add trailing empty string
            List<OptChar> leading = new[] { OptChar.None }.ToList();    // add leading empty string
            return leading.Concat(input).ToArray();
        }

        public string GenerateErrorString_new(string input, Classification c, Random r)
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
                int count = td[key];
                // funny case to handle the fact that FSharpOption.None == null
                var cond_dist = td.Where(kvp => kvp.Key.Item1 == null ? oc == null : kvp.Key.Item1.Equals(oc));
                int total = cond_dist.Aggregate(0, (acc, kvp) => acc + kvp.Value);
                return (double)count / total;
            }).ToArray();

            // calculate the probability of making at least one error
            // might need log-probs here
            double PrTypo = 1.0 - PrsCharNotTypo.Aggregate(1.0, (acc, pr_not_typo) => acc * pr_not_typo);

            // calculate the marginal probabilities of NOT making a
            // transposition for each position in the input
            // note that we do NOT consider the empty strings here
            double[] PrsPosNotTrans = ochars.ToArray().Select((oc, idx) =>
            {
                int count = trd[0];
                int total = trd.Where(kvp => kvp.Key <= input.Length - idx).Select(kvp => kvp.Value).Sum();
                return (double)count / total;
            }).ToArray();

            // calculate the probability of having at least one transposition
            double PrTrans = 1.0 - PrsPosNotTrans.Aggregate(1.0, (acc, pr_not_trans) => acc * pr_not_trans);

            // calculate the relative probability of making a typo vs a transposition
            double RelPrTypo = PrTypo / (PrTypo + PrTrans);

            string output;

            // flip a coin to determine whether our guaranteed error is a typo or a transposition
            if (r.NextDouble() < RelPrTypo)
            {   // is a typo
                // determine the index of the guaranteed typo
                double[] PrsMistype = PrsCharNotTypo.Select(pr => 1.0 - pr).ToArray();
                var i = MultinomialSample(PrsMistype, r);
                // run transposition algorithm & add leading/trailing empty chars
                // we set the guaranteed transposition index to -1 to ensure that no
                // transpositions are guaranteed
                OptChar[] input_t = AddLeadingTrailingSpace(Transposize(ochars, trd, -1, r));
                // run typo algorithm
                output = Typoize(input_t, td, i, r);
            }
            else
            {   // is a transposition
                // determine the index of the guaranteed transposition
                double[] PrsMistype = PrsPosNotTrans.Select(pr => 1.0 - pr).ToArray();
                var i = MultinomialSample(PrsMistype, r);
                // run transposition algorithm & add leading/trailing empty chars
                OptChar[] input_t = AddLeadingTrailingSpace(Transposize(ochars, trd, i, r));
                // run typo algorithm; set guaranteed typo index to -1 to ensure that no
                // type is guaranteed
                output = Typoize(input_t, td, -1, r);
            }

            return input;
        }

        public ErrorString GenerateErrorString(string input, Classification classification)
        {
            //Try to add transposition errors
            Dictionary<int, double> transpositions_distribution = GetDistributionOfTranspositions(classification);
            
            //Keeps track of where characters have ended up after they have been transposed,
            //so that we don't move them more than once
            List<int> transposed_locations = new List<int>();

            string transposed_input = input;
            for (int i = 0; i < transposed_input.Length; i++)
            {
                //if the character in this location has already been transposed, don't transpose it again
                if (transposed_locations.Contains(i))
                {
                    continue;
                }

                char c = transposed_input[i];
                int delta = GetRandomTranspositionFromDistribution(transpositions_distribution);
                int target_index = i + delta;

                //If this target_index doesn't work, randomly select a new one until you find one that works
                //It does not work if it's too large or too small
                while (target_index < 0 || target_index > transposed_input.Length - 1)
                {
                    delta = GetRandomTranspositionFromDistribution(transpositions_distribution);
                    target_index = i + delta;
                }

                //Once we have a delta that works we perform the teleportation. 
                //We have to update all the indices in transposed locations as necessary
                if (delta > 0)
                {
                    List<int> updated_transposed_locations = new List<int>();
                    for (int l = 0; l < transposed_locations.Count; l++)
                    {
                        int transposed_location = transposed_locations[l];
                        //Shift any indices between i and target_index to the left
                        if (transposed_location > i && transposed_location <= target_index)
                        {
                            transposed_location--;
                            updated_transposed_locations.Add(transposed_location);
                        }
                        else
                        {
                            updated_transposed_locations.Add(transposed_location);
                        }
                    }
                    transposed_locations = updated_transposed_locations;
                }
                else if (delta < 0) 
                {
                    List<int> updated_transposed_locations = new List<int>();
                    for (int l = 0; l < transposed_locations.Count; l++)
                    {
                        int transposed_location = transposed_locations[l];
                        //Shift any indices between target_index and i to the right
                        if (transposed_location < i && transposed_location >= target_index)
                        {
                            transposed_location++;
                            updated_transposed_locations.Add(transposed_location);
                        }
                        else
                        {
                            updated_transposed_locations.Add(transposed_location);
                        }
                    }
                    transposed_locations = updated_transposed_locations;
                }
                //Add the error to our error list (only if the delta is non-zero)
                if (delta != 0)
                {
                    //When we have a target index that works, we perform the change
                    transposed_input = transposed_input.Remove(i, 1);
                    transposed_input = transposed_input.Insert(target_index, c + "");

                    LCSError error = LongestCommonSubsequence.Error.NewTranspositionError(i, delta);
                    //And add the indices to the transposed_locations
                    transposed_locations.Add(target_index);
                }

                //If the delta was greater than 0, the next i gets shifted to the left because we remove the current character from its place and move it to the right
                if (delta > 0)
                {
                    i--;
                }
            }

            string[] ti = transposed_input.ToCharArray().Select(c => c.ToString()).ToArray();

            string modified_input = ti.AsParallel().Select((c, i) =>
            {
                char mychar = c[0];

                //If the character in this location has already been transposed, don't introduce typos to it
                if (transposed_locations.Contains(i))
                {
                    return c;
                }

                Dictionary<string, double> distribution = GetDistributionOfStringsForChar(OptChar.Some(mychar), classification);
                return GetRandomStringFromDistribution(distribution);
            }).Aggregate("", (acc, s) => acc + s);

            //Signs and decimals are handled by typo model

            List<LCSError> error_list = new List<LCSError>();
            if (!modified_input.Equals(input))
            {
                // get LCS
                var alignments = LongestCommonSubsequence.LeftAlignedLCS(input, modified_input);
                // find all character additions
                var additions = LongestCommonSubsequence.GetAddedCharIndices(modified_input, alignments);
                // find all character omissions
                var omissions = LongestCommonSubsequence.GetMissingCharIndices(input, alignments);
                // find all transpositions
                var outputs = LongestCommonSubsequence.FixTranspositions(alignments, additions, omissions, input, modified_input);
                // new string
                string modified_input2 = outputs.Item1;
                // new alignments
                var alignments2 = outputs.Item2;
                // new additions
                var additions2 = outputs.Item3;
                // new omissions
                var omissions2 = outputs.Item4;
                // deltas
                var deltas = outputs.Item5;
                // get typos
                var typos = LongestCommonSubsequence.GetTypos(alignments2, input, modified_input2);
                foreach (Tuple<OptChar, string> typo_error in typos)
                {
                    if (typo_error.Item1 != null)
                    {
                        if (!typo_error.Item2.Equals(typo_error.Item1.Value + ""))
                        {
                            LCSError error = LongestCommonSubsequence.Error.NewTypoError(0, typo_error.Item1.Value, typo_error.Item2);
                            error_list.Add(error);
                        }
                    }
                    else
                    {
                        if (!typo_error.Item2.Equals(""))
                        {
                            LCSError error = LongestCommonSubsequence.Error.NewTypoError(0, '\0', typo_error.Item2);
                            error_list.Add(error);
                        }
                    }
                }
                foreach (int delta in deltas)
                {
                    LCSError error = LongestCommonSubsequence.Error.NewTranspositionError(0, delta);
                    error_list.Add(error);
                }
            }
            
            ErrorString output = new Tuple<string, List<LCSError>>(modified_input, error_list);
            
            return output;
        }
    }
}
