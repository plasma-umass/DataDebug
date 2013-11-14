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
        //  Typo
        
        //Dictionaries for the two error types:
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

        public Tuple<int,int> ProcessTypos(string original, string entered)
        {
            // get LCS
            var alignments = LongestCommonSubsequence.LeftAlignedLCS(original, entered);
            // find all character additions
            var additions = LongestCommonSubsequence.GetAddedCharIndices(entered, alignments);
            // find all character omissions
            var omissions = LongestCommonSubsequence.GetMissingCharIndices(original, alignments);
            // find all transpositions
            var outputs = LongestCommonSubsequence.FixTranspositions(alignments, additions, omissions, original, entered);
            // new string
            string entered2 = outputs.Item1;
            // new alignments
            var alignments2 = outputs.Item2;
            // new additions
            var additions2 = outputs.Item3;
            // new omissions
            var omissions2 = outputs.Item4;
            // deltas
            var deltas = outputs.Item5;
            // get typos
            var typos = LongestCommonSubsequence.GetTypos(alignments2, original, entered2);

            int count_deltas = 0;
            int count_typos = 0;

            // train the model for all non-transpositions
            foreach (var alignment in alignments)
            {
                AddTranspositionError(0);
            }

            // train the model for all actual transpositions
            foreach (var delta in deltas)
            {
                AddTranspositionError(delta);
                count_deltas += 1;
            }
            
            // train the model for each "typo", including non-typos
            foreach (var typo in typos)
            {
                OptChar c = typo.Item1;
                string s = typo.Item2;
                AddTypoError(c, s);

                // non-empty string
                if (OptChar.get_IsSome(c))
                {
                    count_typos += c.Value.ToString().Equals(s) ? 0 : 1;
                }
                // empty string
                else
                {
                    count_typos += s.Equals("") ? 0 : 1;
                }
            }
            return new Tuple<int,int>(count_deltas, count_typos);
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

        public Dictionary<Tuple<OptChar,string>,int> GetTypoDict()
        {
            return _typo_dict;
        }

        public Dictionary<int, int> GetTranspositionDict()
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


        //TODO Decide where we are going to keep our trained dictionaries
        public void Serialize()
        {
            string file_name = "C.classification";
            IFormatter formatter = new BinaryFormatter();
            using (Stream stream = new FileStream(file_name, FileMode.Create, FileAccess.Write, FileShare.None))
            {
                formatter.Serialize(stream, this);
            }
        }

        public static Classification Deserialize()
        {
            string file_name = "C:\\Users\\Dimitar Gochev\\Documents\\GitHub\\papers\\DataDebug\\PLDI-2014\\Experiments\\ClassificationData_2013-11-14.bin";
            Classification classification;

            using (Stream stream = new FileStream(file_name, FileMode.Open, FileAccess.Read, FileShare.Read))
            {
                IFormatter formatter = new BinaryFormatter();
                classification = (Classification)formatter.Deserialize(stream);
            }
            return classification;
        }

        public double CharErrorRate()
        {
            // find the total number of typo classifications
            var total = _typo_dict.Select(pair => pair.Value).Sum();

            // find the total number of correct classifications
            var t_correct = _typo_dict.Where(pair =>
            {
                OptChar intended_ochar = pair.Key.Item1;
                string entered_str = pair.Key.Item2;
                // an actual character
                if (OptChar.get_IsSome(intended_ochar))
                {
                    var intended_char = intended_ochar.Value;
                    if (intended_char.ToString().Equals(entered_str))
                    {
                        return true;
                    }
                    else
                    {
                        return false;
                    }
                }
                // this is the empty character
                else
                {
                    if (entered_str.Equals(""))
                    {
                        return true;
                    }
                    else
                    {
                        return false;
                    }
                }
            }).Select(pair => pair.Value).Sum();

            return 1 - (double)t_correct / (double)total;
        }
    }
}
