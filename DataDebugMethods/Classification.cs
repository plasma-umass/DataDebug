using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
using System.Text.RegularExpressions;
using TupleKey = Tuple<char, char, bool>;
using ErrorTypeDict = System.Collections.Generic.Dictionary<TupleKey, int>;
using ErrorTypeDictPair = System.Collections.Generic.KeyValuePair<TupleKey, int>;
namespace DataDebugMethods
{
    class Classification
    {
        //Error types:
        //  Transposition
        //  Digit addition
        //  Digit omission
        //  Sign error (do we want sign addition/ sign omission?)
        //  Typo
        
        //Dictionaries for all error types:
        ErrorTypeDict TranspositionDict = new ErrorTypeDict();
        ErrorTypeDict DigitOmissionDict = new ErrorTypeDict();
        ErrorTypeDict DigitAdditionDict = new ErrorTypeDict();
        ErrorTypeDict SignErrorDict = new ErrorTypeDict();
        ErrorTypeDict DecimalErrorDict = new ErrorTypeDict();
        ErrorTypeDict TypoDict = new ErrorTypeDict();


        //Gets the type of the string input: numeric, letters, or alphanumeric.
        public char getStringType(String s)
        {
            bool allLetters = true;
            foreach (char c in s)
            {
                if (Char.IsDigit(c))
                {
                    allLetters = false;
                }
            }
            if (allLetters)
            {
                return 'L';
            }
            bool allDigits = true;
            foreach (char c in s)
            {
                if (Char.IsLetter(c))
                {
                    allDigits = false;
                }
            }
            if (allDigits)
            {
                return 'N';
            }
            return 'A';
        }

        

    }
}
