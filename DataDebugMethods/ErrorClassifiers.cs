using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace DataDebugMethods
{
    public static class ErrorClassifiers
    {
        public static bool TestMisplacedDecimal(string enteredText, string originalText)
        {
            //Strings must contain at most one decimal point
            int countDecimalPoints = 0;
            for (int i = 0; i < originalText.Length; i++)
            {
                if (originalText[i].Equals('.'))
                {
                    countDecimalPoints++;
                }
            }
            int countDecimalPointsEntered = 0;
            for (int i = 0; i < enteredText.Length; i++)
            {
                if (enteredText[i].Equals('.'))
                {
                    countDecimalPointsEntered++;
                }
            }
            if (countDecimalPoints > 1 || countDecimalPointsEntered > 1)
            {
                //MessageBox.Show("Decimal point error: NO");
                return false;
            }
            //if the strings are different and are the same without decimal points
            if (!originalText.Equals(enteredText) && originalText.Replace(".", "").Equals(enteredText.Replace(".", "")))
            {
                //MessageBox.Show("Decimal point error: YES");
                if (TestDecimalOmission(enteredText, originalText))
                {
                    return false;
                }
                return true;
            }
            else
            {
                //MessageBox.Show("Decimal point error: NO");
                return false;
            }
        } //End TestMisplacedDecimal

        public static bool TestSignOmission(string enteredText, string originalText)
        {
            bool originalStartsWithMinus = true;
            bool remainingCharactersTheSame = true;

            //If the strings are different and original is exactly one character longer than the entered
            if (!originalText.Equals(enteredText) && originalText.Length == (enteredText.Length + 1))
            {
                //Check if first character of original is a "-"
                if (!originalText[0].Equals('-'))
                {
                    originalStartsWithMinus = false;
                }
                //Check if the rest is all the same
                for (int i = 0; i < originalText.Length - 1; i++)
                {
                    if (originalText[i + 1] != enteredText[i])
                    {
                        remainingCharactersTheSame = false;
                    }
                }
            }
            else
            {
                originalStartsWithMinus = false;
                remainingCharactersTheSame = false;
            }

            if (originalStartsWithMinus && remainingCharactersTheSame)
            {
                //MessageBox.Show("Sign omission: YES");
                return true;
            }
            else
            {
                //MessageBox.Show("Sign omission: NO");
                return false;
            }
        } //End TestSignOmission

        public static bool TestDecimalOmission(string enteredText, string originalText)
        {
            //Original string must contain at most one decimal point
            int countDecimalPoints = 0;
            for (int i = 0; i < originalText.Length; i++)
            {
                if (originalText[i].Equals('.'))
                {
                    countDecimalPoints++;
                }
            }
            int countDecimalPointsEntered = 0;
            for (int i = 0; i < enteredText.Length; i++)
            {
                if (enteredText[i].Equals('.'))
                {
                    countDecimalPointsEntered++;
                }
            }
            if (countDecimalPoints > 1 || countDecimalPointsEntered > 1)
            {
                //MessageBox.Show("Decimal point omission: NO");
                return false;
            }
            if (countDecimalPointsEntered < countDecimalPoints && originalText.Replace(".", "").Equals(enteredText.Replace(".", "")))
            {
                //MessageBox.Show("Decimal point omission: YES");
                return true;
            }
            else
            {
                //MessageBox.Show("Decimal point omission: NO");
                return false;
            }
        } //End TestDecimalOmission

        public static bool TestDigitRepeat(string enteredText, string originalText)
        {
            string originalString = originalText;
            string enteredString = enteredText;
            bool startTheSame = false;
            char repeatedChar = 'a';
            //If the strings are different and original string is shorter than entered string
            if (!originalText.Equals(enteredText) && originalText.Length < enteredText.Length)
            {
                //Strings have to start and end with the same characters; only the middle has to be different
                //If the characters at the starting index are the same, remove them
                while (originalString.Length != 0 && originalString[0].Equals(enteredString[0]))
                {
                    //Remove the first characters of originalString and enteredString
                    startTheSame = true;
                    repeatedChar = originalString[0];
                    originalString = originalString.Remove(0, 1);
                    enteredString = enteredString.Remove(0, 1);
                }
                //If the characters at the ending index are the same
                while (originalString.Length != 0 && originalString[originalString.Length - 1].Equals(enteredString[enteredString.Length - 1]))
                {
                    //Remove the last characters of originalString and enteredString
                    originalString = originalString.Remove(originalString.Length - 1);
                    enteredString = enteredString.Remove(enteredString.Length - 1);
                }
                if (!startTheSame)
                {
                    //MessageBox.Show("Digit repeat: NO");
                    //MessageBox.Show("Did not start with the same character, which is required.");
                    return false;
                }
                //Check if the middle part is a single repeated digit which is the same as the digit right before the start of the difference
                //If a digit was repeated, originalString should now be blank (""), and enteredString should have length at least 1
                if (originalString.Length != 0 || enteredString.Length < 1)
                {
                    //MessageBox.Show("Digit repeat: NO");
                    return false;
                }
                //Check if the digit in the part that remains is the same as the one we want
                if (!enteredString[0].Equals(repeatedChar))
                {
                    //MessageBox.Show("Digit repeat: NO");
                    return false;
                }
                //Check if the middle part is composed of the same repeated digit
                for (int i = 0; i < enteredString.Length; i++)
                {
                    if (!enteredString[0].Equals(enteredString[i]))
                    {
                        //MessageBox.Show("Digit repeat: NO");
                        return false;
                    }
                }
                //MessageBox.Show("Digit repeat: YES");
                return true;
            }
            else
            {
                //MessageBox.Show("Digit repeat: NO");
                return false;
            }
        } //End TestDigitRepeat

        public static bool TestExtraDigit(string enteredText, string originalText)
        {
            string originalString = originalText;
            string enteredString = enteredText;
            bool startTheSame = false;
            bool endTheSame = false;
            //If the strings are different and original string is shorter than entered string
            if (!originalText.Equals(enteredText) && originalText.Length < enteredText.Length)
            {
                //Strings have to start and end with the same characters; only the middle has to be different
                //If the characters at the starting index are the same, remove them
                while (originalString.Length != 0 && originalString[0].Equals(enteredString[0]))
                {
                    //Remove the first characters of originalString and enteredString
                    startTheSame = true;
                    originalString = originalString.Remove(0, 1);
                    enteredString = enteredString.Remove(0, 1);
                }
                //If the characters at the ending index are the same
                while (originalString.Length != 0 && originalString[originalString.Length - 1].Equals(enteredString[enteredString.Length - 1]))
                {
                    //Remove the last characters of originalString and enteredString
                    endTheSame = true;
                    originalString = originalString.Remove(originalString.Length - 1);
                    enteredString = enteredString.Remove(enteredString.Length - 1);
                }
                if (!(startTheSame || endTheSame))
                {
                    return false;
                }
                //Check if the middle part is a single repeated digit which is the same as the digit right before the start of the difference
                //If a digit was repeated, originalString should now be blank (""), and enteredString should have length at least 1
                if (originalString.Length != 0 || enteredString.Length < 1)
                {
                    //MessageBox.Show("Digit repeat: NO");
                    return false;
                }
                //Check if the middle part is composed of the same repeated digit
                for (int i = 0; i < enteredString.Length; i++)
                {
                    if (!enteredString[0].Equals(enteredString[i]))
                    {
                        return false;
                    }
                }
                if (TestDigitRepeat(enteredText, originalText))
                {
                    return false;
                }
                return true;
            }
            else
            {
                return false;
            }
        } //End TestExtraDigit

        public static bool TestWrongDigit(string enteredText, string originalText)
        {
            string originalString = originalText;
            string enteredString = enteredText;
            bool startTheSame = false;
            bool endTheSame = false;
            //If the strings are different and original string is shorter than entered string
            if (!originalText.Equals(enteredText) && originalText.Length == enteredText.Length)
            {
                //Strings have to start and end with the same characters; only the middle has to be different
                //If the characters at the starting index are the same, remove them
                while (enteredString.Length != 0 && originalString[0].Equals(enteredString[0]))
                {
                    //Remove the first characters of originalString and enteredString
                    startTheSame = true;
                    originalString = originalString.Remove(0, 1);
                    enteredString = enteredString.Remove(0, 1);
                }
                //If the characters at the ending index are the same
                while (enteredString.Length != 0 && originalString[originalString.Length - 1].Equals(enteredString[enteredString.Length - 1]))
                {
                    //Remove the last characters of originalString and enteredString
                    endTheSame = true;
                    originalString = originalString.Remove(originalString.Length - 1);
                    enteredString = enteredString.Remove(enteredString.Length - 1);
                }
                //They have to either start the same or end the same
                if (!(startTheSame || endTheSame))
                {
                    return false;
                }
                //Check if the middle part is a single repeated digit which is the same as the digit right before the start of the difference
                //If a digit was repeated, originalString should now be blank (""), and enteredString should have length at least 1
                if (originalString.Length == 1 && enteredString.Length == 1)
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
            else
            {
                return false;
            }
        } //End TestWrongDigit

        public static bool TestDigitOmission(string enteredText, string originalText)
        {
            string originalString = originalText;
            //If there is a sign omission, we want to compare the numeric part only
            if (TestSignOmission(enteredText, originalText))
            {
                originalString = originalString.Substring(1);
            }
            //If there is a decimal omission, we want to compare the numeric part only
            if (TestDecimalOmission(enteredText, originalText))
            {
                originalString = originalString.Replace(".", "");
            }
            string enteredString = enteredText;
            bool startTheSame = false;
            bool endTheSame = false;
            //If the strings are different and original string is shorter than entered string
            if (!originalText.Equals(enteredText) && originalText.Length > enteredText.Length)
            {
                //Strings have to start and end with the same characters; only the middle has to be different
                //If the characters at the starting index are the same, remove them
                while (enteredString.Length != 0 && originalString[0].Equals(enteredString[0]))
                {
                    //Remove the first characters of originalString and enteredString
                    startTheSame = true;
                    originalString = originalString.Remove(0, 1);
                    enteredString = enteredString.Remove(0, 1);
                }
                //If the characters at the ending index are the same
                while (enteredString.Length != 0 && originalString[originalString.Length - 1].Equals(enteredString[enteredString.Length - 1]))
                {
                    //Remove the last characters of originalString and enteredString
                    endTheSame = true;
                    originalString = originalString.Remove(originalString.Length - 1);
                    enteredString = enteredString.Remove(enteredString.Length - 1);
                }
                //They have to either start the same or end the same
                if (!(startTheSame || endTheSame))
                {
                    return false;
                }
                //Check if the middle part is a single repeated digit which is the same as the digit right before the start of the difference
                //If a digit was repeated, originalString should now be blank (""), and enteredString should have length at least 1
                if (originalString.Length == 1 && enteredString.Length == 0)
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
            else
            {
                return false;
            }
        } //End TestDigitOmission

        public static bool TestBlank(string enteredText, string originalText)
        {
            if (enteredText.Length == 0 && originalText.Length > 0)
            {
                return true;
            }
            else
            {
                return false;
            }
        } //End TestBlank

        public static bool TestDigitTransposition(string enteredText, string originalText)
        {
            //For a transposition to occur, the original text has to have at least two characters. Also, the original text and entered text must differ, but have the same length.
            //We assume there is only one transposition in the number -- that is, there must be one transposition, and all other characters must be correct; otherwise this will count as a general typo. 
            if (originalText.Length >= 2 && !enteredText.Equals(originalText) && enteredText.Length == originalText.Length)
            {
                //Look for transpositions at each index in the string
                for (int i = 0; i < originalText.Length - 1; i++)
                {
                    //For a transposition to exist, two consecutive characters in the original string (which are different) must be transposed (typed in the reverse order) in the entered string. 
                    if (!originalText[i].Equals(originalText[i + 1]) && originalText[i].Equals(enteredText[i + 1]) && originalText[i + 1].Equals(enteredText[i]))
                    {
                        //Once we have identified a transposition, we have to check if all other characters besides the ones at index i and i+1 are correct
                        bool othersMatch = true;
                        for (int j = 0; j < originalText.Length; j++)
                        {
                            if (j == i || j == i + 1)
                            {
                                continue;
                            }
                            if (originalText[j].Equals(enteredText[j]) == false)
                            {
                                othersMatch = false;
                            }
                        }
                        if (othersMatch == true)
                        {
                            return true;
                        }
                        else
                        {
                            return false;
                        }
                    }
                }
                return false;
            }
            else
            {
                return false;
            }
        } //End TestDigitTransposition

        public static bool TestSignError(string enteredText, string originalText)
        {
            //Both strings have to have at least one character besides a - sign
            if (enteredText.Replace("-", "").Length >= 1 && originalText.Replace("-", "").Length >= 1)
            {
                //If the first character in the original string is a '-', but it is not in the entered string, this is a sign error
                if (originalText[0].Equals('-') && !enteredText[0].Equals('-'))
                {
                    return true;
                }
                //If the first character in the entered string is a '-', but it is not in the original string, this is a sign error
                else if (!originalText[0].Equals('-') && enteredText[0].Equals('-'))
                {
                    return true;
                }
                //If there is no difference in the signs, this is not a sign error
                else
                {
                    return false;
                }
            }
            //If the strings are shorter than one character (ignoring - signs), there is no sign error
            else
            {
                return false;
            }
        }  //End TestSignError
    }
}
