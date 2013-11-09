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

        private Dictionary<string, double> GenerateDistributionForChar(OptChar c, Classification classification)
        {
            var typo_dict = classification.GetTypoDict();
            var kvps = typo_dict.Where(pair => pair.Key.Item1.Equals(c));
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
            
            //Keeps track of where things have ended up after they have been transposed,
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
                int swap_index = i + delta;

                //If this swap_index doesn't work, randomly select a new one until you find one that works
                //  It might not work for any of these three reasons: 
                //      1. It's already been transposed
                //      2. It's too large
                //      3. It's too small
                while (transposed_locations.Contains(swap_index) || swap_index < 0 || swap_index > transposed_input.Length - 1)
                {
                    delta = GetRandomTranspositionFromDistribution(transpositions_distribution);
                    swap_index = i + delta;
                }

                //When we have a swap index that works, we perform the swap
                char swap_char = transposed_input[swap_index];
                transposed_input = transposed_input.Remove(i, 1);
                transposed_input = transposed_input.Insert(i, swap_char + "");
                transposed_input = transposed_input.Remove(swap_index, 1);
                transposed_input = transposed_input.Insert(swap_index, c + "");
                
                //Add the error to our error list (only if the delta is non-zero)
                if (delta != 0)
                {
                    LCSError error = LongestCommonSubsequence.Error.NewTranspositionError(i, delta);
                    error_list.Add(error);
                    //And add the indices to the transposed_locations
                    transposed_locations.Add(i);
                    transposed_locations.Add(swap_index);
                }
            }

            String modified_input = "";
            //Try to add typo errors
            for (int i = 0; i < transposed_input.Length; i++)
            {
                char c = transposed_input[i];
                
                //If the character in this location has already been transposed, don't introduce typos to it
                if (transposed_locations.Contains(i))
                {
                    modified_input += "" + c;
                    continue;
                }

                Dictionary<string, double> distribution = GetDistributionOfStringsForChar(OptChar.Some(c), classification);
                string str = GetRandomStringFromDistribution(distribution);

                //If this was an error, add it to the error list
                if (!str.Equals("" + c))
                {
                    LCSError error = LongestCommonSubsequence.Error.NewTypoError(i, c, str);
                    error_list.Add(error);
                }
                modified_input += str;
            }

            //Signs and decimals are handled by typo model

            ErrorString output = new Tuple<string, List<LCSError>>(modified_input, error_list);
            
            return output;
        }
    }
}
