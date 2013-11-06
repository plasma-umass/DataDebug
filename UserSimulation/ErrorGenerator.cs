using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.FSharp.Core;
using Sign = LongestCommonSubsequence.Sign;
using LCSError = LongestCommonSubsequence.Error;
using ErrorString = Tuple<string, List<LCSError>>;

namespace UserSimulation
{
    class ErrorGenerator
    {
        public static ErrorString GenerateErrorString(string input)
        {
            List<LCSError> error_list = new List<LCSError>();
            String modified_input = input;
            //try to add a sign error
            Sign s = Classification.getSign(input);
            var kvps = _sign_dict.Where(pair => pair.Key.Item1 == s);
            var sum = kvps.Select(pair => pair.Value).Sum();
            var rng = new Random();
            var bins = kvps.Select(pair => Enumerable.Repeat(pair.Key,pair.Value)).SelectMany( i => i);
            var tpl = bins.ElementAt(rng.Next(bins.Count()));
            Sign s2 = tpl.Item2;
            
            if (s != s2)
            {
                LCSError error = new LCSError(s, s2);
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
