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
        //private _typo_array;
        //private _transposition_array;
        //private _sign_array;
        private Dictionary<OptChar, Dictionary<string,double>> _char_distributions_dict = new Dictionary<OptChar,Dictionary<string,double>>();

        private Dictionary<int, double> _transpositions_distribution_dict = new Dictionary<int, double>();

        private Dictionary<string, double> GetDistributionOfStringsForChar(OptChar c, Classification classification)
        {
            OptChar key = c;
            Dictionary<string, double> distribution;
            distribution = GenerateDistributionForChar(key, classification);
            //If our dictionary does not have any information about this character, we return the character with probability 1.0
            if (distribution.Count == 0)
            {
                distribution.Add("" + c.Value, 1.0);
            }
            return distribution;
        }

        private Dictionary<string, double> GetDistributionOfStringsForChaz(OptChar c, Classification classification)
        {
            OptChar key = c;
            Dictionary<string, double> distribution;
            //if we have already generated a distribution for this character, return it
            if (_char_distributions_dict.TryGetValue(key, out distribution))
            {
                return distribution;
            }
            else //otherwise generate the distribution and then return it
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
            //var distribution = kvps.Select(pair => Enumerable.Repeat(pair.Key, pair.Value)).SelectMany(i => i);
            return distribution.ToDictionary(pair => pair.Key, pair => pair.Value);
        }

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
            //var kvps = transposition_dict.Where(pair => pair.Key.Equals(i));
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

        public ErrorString GenerateErrorString(string input, Classification classification)
        {
            List<LCSError> error_list = new List<LCSError>();
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
                    error_list.Add(error);
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

            //TODO get rid of error_list; instead, at the end of GenerateErrors, check if the input 
            //  was changed at all; if yes, run the classifier to find out how it changed
            if (!modified_input.Equals(input))
            {
                //error_list = get errors from Classification

                error_list.Clear();
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
