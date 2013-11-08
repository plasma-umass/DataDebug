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
        private Dictionary<Sign, Dictionary<Sign,double>> _sign_distributions_dict = new Dictionary<Sign,Dictionary<Sign,double>>();

        private Dictionary<OptChar, Dictionary<string,double>> _char_distributions_dict = new Dictionary<OptChar,Dictionary<string,double>>();

        public Dictionary<string, double> GetDistributionForChar(OptChar c, Classification classification)
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

        public Dictionary<Sign, double> GetDistributionForSign(Sign s, Classification classification)
        {
            Sign key = s;
            Dictionary<Sign, double> distribution;
            if (_sign_distributions_dict.TryGetValue(key, out distribution))
            {
                return distribution; 
            }
            else
            {
                distribution = GenerateDistributionForSign(key, classification);
                if (distribution.Count == 0)
                {
                    distribution.Add(s, 1.0);
                }
                _sign_distributions_dict.Add(key, distribution);
                return distribution;
            }
        }

        public Dictionary<string, double> GenerateDistributionForChar(OptChar c, Classification classification)
        {
            var typo_dict = classification.GetTypoDict();
            var kvps = typo_dict.Where(pair => pair.Key.Item1.Equals(c));
            var sum = kvps.Select(pair => pair.Value).Sum();
            var distribution = kvps.Select(pair => new KeyValuePair<string,double>(pair.Key.Item2, (double) pair.Value / sum));
            //var distribution = kvps.Select(pair => Enumerable.Repeat(pair.Key, pair.Value)).SelectMany(i => i);
            return distribution.ToDictionary(pair => pair.Key, pair => pair.Value);
        }

        public Dictionary<Sign,double> GenerateDistributionForSign(Sign s, Classification classification)
        {
            var sign_dict = classification.GetSignDict();
            var kvps = sign_dict.Where(pair => pair.Key.Item1 == s);
            var sum = kvps.Select(pair => pair.Value).Sum();
            var distribution = kvps.Select(pair => new KeyValuePair<Sign,double>(pair.Key.Item2, (double) pair.Value / sum));
            //var distribution = kvps.Select(pair => Enumerable.Repeat(pair.Key, pair.Value)).SelectMany(i => i);
            return distribution.ToDictionary(pair => pair.Key, pair => pair.Value);
        }

        public Sign GetRandomSignFromDistribution(Dictionary<Sign, double> distribution)
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

        public string GetRandomStringFromDistribution(Dictionary<string, double> distribution)
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

        public ErrorString GenerateErrorString(string input, Classification classification)
        {
            List<LCSError> error_list = new List<LCSError>();
            String modified_input = "";
            //try to add a sign error
            Sign s = Classification.GetSign(input);
            Dictionary<Sign,double> distribution = GetDistributionForSign(s, classification);

            Sign s2 = GetRandomSignFromDistribution(distribution);

            for (int i = 0; i < input.Length; i++)
            {
                char c = input[i];
                Dictionary<string, double> distribution2 = GetDistributionForChar(OptChar.Some(c), classification);
                string str = GetRandomStringFromDistribution(distribution2);
                modified_input += str;
            }

            //TODO if we have a character that we don't have as a key in our dictionary already, we should just return that character
                        
            if (s != s2)
            {
                LCSError error = LongestCommonSubsequence.Error.NewSignError(s, s2);
                error_list.Add(error);
                if (s == Sign.Empty)
                {
                    if (s2 == Sign.Plus)
                    {
                        modified_input = "+" + modified_input;
                    }
                    else if (s2 == Sign.Minus)
                    {
                        modified_input = "-" + modified_input;
                    }
                }
                else
                {
                    if (s2 == Sign.Plus)
                    {
                        modified_input = "+" + modified_input.Remove(0,1);
                    }
                    else if (s2 == Sign.Minus)
                    {
                        modified_input = "-" + modified_input.Remove(0, 1);
                    }
                    else
                    {
                        modified_input = modified_input.Remove(0, 1);
                    }
                }
            }            
            
            //Decimals are handled by typo model

            ErrorString output = new Tuple<string, List<LCSError>>(modified_input, error_list);
            
            return output;
        }
    }
}
