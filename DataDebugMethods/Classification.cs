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

namespace DataDebugMethods
{
    public enum Sign
    {
        Plus,
        Minus,
        Empty
    }


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
        private Dictionary<Tuple<Sign,Sign>,int> _sign_dict = new Dictionary<Tuple<Sign,Sign>,int>();
        // key: <decimal location>, value: frequency count
        private Dictionary<int, int> _decimal_misplacement_dict = new Dictionary<int, int>();
        // key: <decimal is missing>, value: frequency count
        private Dictionary<bool, int> _decimal_missing_dict = new Dictionary<bool, int>();
        //ErrorTypeDict TranspositionDict = new ErrorTypeDict();
        //ErrorTypeDict DigitOmissionDict = new ErrorTypeDict();
        //ErrorTypeDict DigitAdditionDict = new ErrorTypeDict();
        //ErrorTypeDict DecimalErrorDict = new ErrorTypeDict();
        //ErrorTypeDict TypoDict = new ErrorTypeDict();

        public void AddSignError(Sign correct, Sign entered)
        {
            var key = new Tuple<Sign,Sign>(correct,entered);
            int value;
            if (_sign_dict.TryGetValue(key, out value)) {
                _sign_dict[key] += 1;
            } else {
                _sign_dict.Add(key, 1);
            }
        }

        public void AddDecimalOmission(bool is_missing)
        {
            int value;
            if (_decimal_missing_dict.TryGetValue(is_missing, out value)) {
                _decimal_missing_dict[is_missing] += 1;
            } else {
                _decimal_missing_dict.Add(is_missing, 1);
            }
        }

        public void AddDecimalMisplacement(int delta)
        {
            int value;
            if (_decimal_misplacement_dict.TryGetValue(delta, out value)) {
                _decimal_misplacement_dict[delta] += 1;
            } else {
                _decimal_misplacement_dict.Add(delta, 1);
            }
        }

        public void ProcessTypos(string original, string entered)
        {
            bool has_errors = true;
            string entered_mod = entered;
            while (has_errors)
            {
                OptString fix = OptString.None;
                has_errors = false;

                // look for a sign error
                fix = HasSignError(original, entered_mod);
                if (fix != OptString.None)
                {
                    entered_mod = fix.Value;
                    has_errors = true;
                }

                // next test here
            }
        }

        public OptString HasSignError(string original, string entered)
        {
            // sign for orig
            Sign orig_sign;
            if (original[0] == '+') 
            {
                orig_sign = Sign.Plus; 
            }
            else if (original[0] == '-')
            {
                orig_sign = Sign.Minus;
            }
            else
            {
                orig_sign = Sign.Empty;
            }

            //If the entered string is blank, return empty optstring
            if (entered.Length < 1)
            {
                AddSignError(orig_sign, Sign.Empty);
                return OptString.None;
            }

            // sign for entered
            Sign ent_sign;
            if (entered[0] == '+')
            {
                ent_sign = Sign.Plus;
            }
            else if (entered[0] == '-')
            {
                ent_sign = Sign.Minus;
            }
            else
            {
                ent_sign = Sign.Empty;
            }

            // update probabilities
            AddSignError(orig_sign, ent_sign);

            // look at the first characters
            var fc_orig = original[0];
            var fc_ent = entered[0];

            // does the original string have a sign?
            var orig_has_sign = false;
            if (fc_orig == '-' || fc_orig == '+')
            {
                orig_has_sign = true;
            }

            // does the entered string have a sign?
            var ent_has_sign = false;
            if (fc_ent == '-' || fc_ent == '+')
            {
                ent_has_sign = true;
            }

            // if the original string had no sign but the entered one did
            // erase the sign in the entered string
            if (ent_has_sign && !orig_has_sign)
            {
                return new OptString(entered.Remove(0, 1));
            }

            // if the original string had a sign but the entered string
            // did not, remove the sign in the entered string
            if (!ent_has_sign && orig_has_sign)
            {
                return new OptString(fc_orig + entered);
            }

            // both have signs but are not the same
            if (ent_has_sign && orig_has_sign && (orig_sign != ent_sign))
            {
                return new OptString(fc_orig + entered.Remove(0,1));
            }

            // no sign errors
            return OptString.None;
        }

        public OptString TestMisplacedDecimal(string original, string entered)
        {
            // original must contain at most one decimal point
            int countDecimalPoints = 0;
            for (int i = 0; i < original.Length; i++)
            {
                if (original[i].Equals('.'))
                {
                    countDecimalPoints++;
                }
            }
            if (countDecimalPoints != 1 || entered.LastIndexOf('-') == -1)
            {
                AddDecimalMisplacement(0);
                return OptString.None;
            }

            // index of decimal
            var orig_idx = original.LastIndexOf('.');

            // if the entered string is not as long as the split original's lhs, bail
            if (entered.Length < orig_idx)
            {
                AddDecimalMisplacement(0);
                return OptString.None;
            }

            // split the error string by the original index
            var ent_lhs = entered.Substring(0, orig_idx);
            var ent_rhs = entered.Substring(orig_idx);

            // find the first occurence of a decimal for each side
            var pos_lhs = ent_lhs.LastIndexOf('.');
            var pos_rhs = ent_rhs.IndexOf('.');


            // the order matters here... BE CAREFUL!
            // there is no decimal on the left
            if (pos_lhs == -1)
            {
                AddDecimalMisplacement(pos_rhs);
                return OptString.Some(entered.Remove(orig_idx + pos_rhs).Insert(orig_idx, "."));
            }
            // there is no decimal on the right
            if (pos_rhs == -1)
            {
                AddDecimalMisplacement(-pos_lhs);
                return OptString.Some(entered.Insert(orig_idx, ".").Remove(orig_idx - pos_lhs));
            }
            // there are decimals on both sides, but the left side is closer
            if (pos_lhs < pos_rhs)
            {
                AddDecimalMisplacement(-pos_lhs);
                return OptString.Some(entered.Insert(orig_idx, ".").Remove(orig_idx - pos_lhs));
            }
            // there are decimals on both sides, but the right side is closer
            else
            {
                AddDecimalMisplacement(pos_rhs);
                return OptString.Some(entered.Remove(orig_idx + pos_rhs).Insert(orig_idx, "."));
            }
        } //End TestMisplacedDecimal

        public OptString TestDecimalOmission(string entered, string original)
        {
            // Original string must contain at most one decimal point
            int countDecimalPoints = 0;
            int decimal_index = 0;
            for (int i = 0; i < original.Length; i++)
            {
                if (original[i].Equals('.'))
                {
                    countDecimalPoints++;
                    decimal_index = i;
                }
            }
            // if there's more than one decimal in the entered string, it is not a number
            // or if the entered string contains decimals, this isn't a decimal omission
            // or if the entered string is shorter than the location of the decimal in the original string (we can't add the decimal in that case)
            // so we don't care
            if (countDecimalPoints != 1 || entered.LastIndexOf('.') != -1 || entered.Length <= decimal_index)
            {
                AddDecimalOmission(false);
                return OptString.None;
            }

            AddDecimalOmission(true);
            
            return OptString.Some(entered.Insert(decimal_index, "."));
        } //End TestDecimalOmission


        //public static bool TestDigitTransposition(string enteredText, string originalText)
        //{
        //    //For a transposition to occur, the original text has to have at least two characters. Also, the original text and entered text must differ, but have the same length.
        //    //We assume there is only one transposition in the number -- that is, there must be one transposition, and all other characters must be correct; otherwise this will count as a general typo. 
        //    if (originalText.Length >= 2 && !enteredText.Equals(originalText) && enteredText.Length == originalText.Length)
        //    {
        //        //Look for transpositions at each index in the string
        //        for (int i = 0; i < originalText.Length - 1; i++)
        //        {
        //            //For a transposition to exist, two consecutive characters in the original string (which are different) must be transposed (typed in the reverse order) in the entered string. 
        //            if (!originalText[i].Equals(originalText[i + 1]) && originalText[i].Equals(enteredText[i + 1]) && originalText[i + 1].Equals(enteredText[i]))
        //            {
        //                //Once we have identified a transposition, we have to check if all other characters besides the ones at index i and i+1 are correct
        //                bool othersMatch = true;
        //                for (int j = 0; j < originalText.Length; j++)
        //                {
        //                    if (j == i || j == i + 1)
        //                    {
        //                        continue;
        //                    }
        //                    if (originalText[j].Equals(enteredText[j]) == false)
        //                    {
        //                        othersMatch = false;
        //                    }
        //                }
        //                if (othersMatch == true)
        //                {
        //                    return true;
        //                }
        //                else
        //                {
        //                    return false;
        //                }
        //            }
        //        }
        //        return false;
        //    }
        //    else
        //    {
        //        return false;
        //    }
        //} //End TestDigitTransposition


        ////Gets the type of the string input: numeric, letters, or alphanumeric.
        //public char getStringType(String s)
        //{
        //    bool allLetters = true;
        //    foreach (char c in s)
        //    {
        //        if (Char.IsDigit(c))
        //        {
        //            allLetters = false;
        //        }
        //    }
        //    if (allLetters)
        //    {
        //        return 'L';
        //    }
        //    bool allDigits = true;
        //    foreach (char c in s)
        //    {
        //        if (Char.IsLetter(c))
        //        {
        //            allDigits = false;
        //        }
        //    }
        //    if (allDigits)
        //    {
        //        return 'N';
        //    }
        //    return 'A';
        //}

        

    }
}
