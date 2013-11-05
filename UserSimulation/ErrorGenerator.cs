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
            ErrorString output = new Tuple<string, List<LCSError>>("",null);
            //try to add a sign error

            //

            //Adding a decimal is handled by inserted characters
            return output;
        }
    }
}
