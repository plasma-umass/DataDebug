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
            String modified_input = "";
            //Try to add transposition errors
            Dictionary<int, double> transpositions_distribution = GetDistributionOfTranspositions(classification);
            
            //Keeps track of where things have ended up after they have been transposed,
            //so that we don't move them more than once
            List<int> dont_transpose = new List<int>();
            
            for (int i = 0; i < input.Length; i++)
            {
                //if this location has already been transposed, don't transpose it again
                if (dont_transpose.Contains(i))
                {
                    continue;
                }

                char c = input[i];
                int delta = GetRandomTranspositionFromDistribution(transpositions_distribution);

                //keeps track of the delta after any necessary truncations
                int actual_delta = delta;
                //if the delta is too large, truncate it
                if (i + delta >= input.Length)
                {
                    actual_delta = input.Length - 1 - i;
                }
                //if the delta is too small, truncate it
                else if (i + delta < 0)
                {
                    actual_delta = -i;
                }

                //If this was an error, add it to the error list
                if (actual_delta != 0)
                {
                    LCSError error = LongestCommonSubsequence.Error.NewTranspositionError(i, actual_delta);
                    error_list.Add(error);
                }
                //If actual_delta == 0
                else
                {
                    continue;
                }

                //if we are trying to move a character to the very end of the string
                if (i + actual_delta == input.Length - 1)
                {
                    input = input.Remove(i, 1);
                    input = input + c;
                    dont_transpose.Add(input.Length - 1);
                }
                else
                {
                    input = input.Remove(i, 1);
                    input = input.Insert(i + actual_delta, c + "");
                    dont_transpose.Add(i + actual_delta);
                }
            }

            //Try to add typo errors
            for (int i = 0; i < input.Length; i++)
            {
                char c = input[i];
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
