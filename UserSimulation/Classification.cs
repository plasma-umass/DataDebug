using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
using System.Text.RegularExpressions;
//using TupleKey = Tuple<char, char, bool>;
//using ErrorTypeDict = System.Collections.Generic.Dictionary<TupleKey, int>;
//using ErrorTypeDictPair = System.Collections.Generic.KeyValuePair<TupleKey, int>;
using Microsoft.FSharp.Core;
using OptString = Microsoft.FSharp.Core.FSharpOption<string>;
using OptInt = Microsoft.FSharp.Core.FSharpOption<int>;
using OptChar = Microsoft.FSharp.Core.FSharpOption<char>;
using Sign = LongestCommonSubsequence.Sign;
using System.IO;
using System.Runtime.Serialization.Formatters.Binary;
using System.Runtime.Serialization;

namespace UserSimulation
{
    [Serializable]
    public class Classification
    {
        //Error types:
        //  Transposition
        //  Digit addition
        //  Digit omission
        //  Sign error (do we want sign addition/ sign omission?)
        //  Decimal addition, omission, misplaced
        //  Typo
        
        //Dictionaries for all error types:
        // key: <correct sign, entered sign>, value: frequency count
        //private Dictionary<Tuple<Sign,Sign>,int> _sign_dict = new Dictionary<Tuple<Sign,Sign>,int>();
        // key: <char that was supposed to be typed, string that was typed>, value: frequency count
        private Dictionary<Tuple<OptChar, string>, int> _typo_dict = new Dictionary<Tuple<OptChar, string>, int>();
        // key: Delta (difference from original location; 0 if there wasn't a transposition), value: frequency count
        private Dictionary<int, int> _transposition_dict = new Dictionary<int, int>();

        public void AddTypoError(OptChar intended, string entered)
        {
            var key = new Tuple<OptChar, string>(intended, entered);
            int value; 
            if (_typo_dict.TryGetValue(key, out value)) 
            {
                _typo_dict[key] += 1;
            }
            else 
            {
                _typo_dict.Add(key, 1);
            }
        }

        public void AddTranspositionError(int delta)
        {
            var key = delta;
            int value;
            if (_transposition_dict.TryGetValue(key, out value))
            {
                _transposition_dict[key] += 1;
            }
            else
            {
                _transposition_dict.Add(key, 1);
            }
        }

        //public void AddSignError(Sign correct, Sign entered)
        //{
        //    var key = new Tuple<Sign,Sign>(correct,entered);
        //    int value;
        //    if (_sign_dict.TryGetValue(key, out value))
        //    {
        //        _sign_dict[key] += 1;
        //    }
        //    else
        //    {
        //        _sign_dict.Add(key, 1);
        //    }
        //}

        //public void AddDecimalOmission()
        //{
        //    int value;
        //    if (_decimal_misplacement_dict.TryGetValue(OptInt.None, out value))
        //    {
        //        _decimal_misplacement_dict[OptInt.None] += 1;
        //    } else {
        //        _decimal_misplacement_dict.Add(OptInt.None, 1);
        //    }
        //}

        //public void AddDecimalMisplacement(OptInt delta)
        //{
        //    int value;
        //    if (_decimal_misplacement_dict.TryGetValue(delta, out value)) {
        //        _decimal_misplacement_dict[delta] += 1;
        //    } else {
        //        _decimal_misplacement_dict.Add(delta, 1);
        //    }
        //}

        public void ProcessTypos(string original, string entered)
        {
            // get LCS
            var alignments = LongestCommonSubsequence.LeftAlignedLCS(original, entered);
            // find all character additions
            var additions = LongestCommonSubsequence.GetAddedCharIndices(entered, alignments);
            // find all character omissions
            var omissions = LongestCommonSubsequence.GetMissingCharIndices(original, alignments);
            // find all transpositions
            var transpositions = LongestCommonSubsequence.FixTranspositions(alignments, additions, omissions, original, entered);
            // remove all transpositions from alignment list
            //var additions2 = additions.Where(a => !transpositions.Select(tpair => tpair.Item1).Contains(a));
            //var omissions2 = omissions.Where(o => !transpositions.Select(tpair => tpair.Item2).Contains(o));
            // remember: alignments is a list of (original position, entered position) pairs
            //var alignments2 = alignments.Where(apair => 
            // get typos
            var typos = LongestCommonSubsequence.GetTypos(alignments, original, entered);
            // now train the classifier for each remaining error
            //foreach (var tpair in transpositions)
            //{
            //    // calculate delta = addition position - omission position
            //    var delta = tpair.Item1 - tpair.Item2;

            //    // update probability
            //    AddTranspositionError(delta);
            //}
            //foreach (int addpos in additions2)
            //{

            //}
        }

        public static Sign GetSign(string input)
        {
            Sign s;
            if (input.Length < 1)
            {
                s = Sign.Empty;
                return s;
            }
            if (input[0] == '+')
            {
                s = Sign.Plus;
            }
            else if (input[0] == '-')
            {
                s = Sign.Minus;
            }
            else
            {
                s = Sign.Empty;
            }
            return s;
        }

        //public OptString HasSignError(string original, string entered)
        //{
        //    // sign for orig
        //    Sign orig_sign = GetSign(original);
          
        //    //If the entered string is blank, return empty optstring
        //    if (entered.Length < 1)
        //    {
        //        AddSignError(orig_sign, Sign.Empty);
        //        return OptString.None;
        //    }

        //    // sign for entered
        //    Sign ent_sign = GetSign(entered);
            
        //    // update probabilities
        //    AddSignError(orig_sign, ent_sign);

        //    // look at the first characters
        //    var fc_orig = original[0];
        //    var fc_ent = entered[0];

        //    // does the original string have a sign?
        //    var orig_has_sign = false;
        //    if (fc_orig == '-' || fc_orig == '+')
        //    {
        //        orig_has_sign = true;
        //    }

        //    // does the entered string have a sign?
        //    var ent_has_sign = false;
        //    if (fc_ent == '-' || fc_ent == '+')
        //    {
        //        ent_has_sign = true;
        //    }

        //    // if the original string had no sign but the entered one did
        //    // erase the sign in the entered string
        //    if (ent_has_sign && !orig_has_sign)
        //    {
        //        return new OptString(entered.Remove(0, 1));
        //    }

        //    // if the original string had a sign but the entered string
        //    // did not, remove the sign in the entered string
        //    if (!ent_has_sign && orig_has_sign)
        //    {
        //        return new OptString(fc_orig + entered);
        //    }

        //    // both have signs but are not the same
        //    if (ent_has_sign && orig_has_sign && (orig_sign != ent_sign))
        //    {
        //        return new OptString(fc_orig + entered.Remove(0,1));
        //    }

        //    // no sign errors
        //    return OptString.None;
        //}

        internal Dictionary<Tuple<OptChar,string>,int> GetTypoDict()
        {
            return _typo_dict;
        }

        internal Dictionary<int, int> GetTranspositionDict()
        {
            return _transposition_dict;
        }

        public void SetTypoDict(Dictionary<Tuple<OptChar, string>, int> dict)
        {
            _typo_dict = dict;
        }

        public void SetTranspositionDict(Dictionary<int, int> dict)
        {
            _transposition_dict = dict;
        }

        public void Serialize(string file_name)
        {
            IFormatter formatter = new BinaryFormatter();
            using (Stream stream = new FileStream(file_name, FileMode.Create, FileAccess.Write, FileShare.None))
            {
                formatter.Serialize(stream, this);
            }
        }

        public static Classification Deserialize(string file_name)
        {
            Classification classification;

            using (Stream stream = new FileStream(file_name, FileMode.Open, FileAccess.Read, FileShare.Read))
            {
                // deserialize
                IFormatter formatter = new BinaryFormatter();
                classification = (Classification)formatter.Deserialize(stream);
            }
            return classification;
        }
    }
}
