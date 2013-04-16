using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using ErrorClassifiers = DataDebugMethods.ErrorClassifiers;

namespace ErrorClassifier
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void signOmission_Click(object sender, EventArgs e)
        {
            if (ErrorClassifiers.TestSignOmission(enteredText.Text, originalText.Text))
            {
                signOmissionLabel.ForeColor = System.Drawing.Color.Green;
                signOmissionLabel.Text = " Y ";
            }
            else
            {
                signOmissionLabel.ForeColor = System.Drawing.Color.Red;
                signOmissionLabel.Text = " N ";
            }
        }

        private void decimalPoint_Click(object sender, EventArgs e)
        {
            if (ErrorClassifiers.TestMisplacedDecimal(enteredText.Text, originalText.Text))
            {
                misplacedDecimalLabel.ForeColor = System.Drawing.Color.Green;
                misplacedDecimalLabel.Text = " Y ";
            }
            else
            {
                misplacedDecimalLabel.ForeColor = System.Drawing.Color.Red;
                misplacedDecimalLabel.Text = " N ";
            }   
        }

        private void digitRepeat_Click(object sender, EventArgs e)
        {
            if (ErrorClassifiers.TestDigitRepeat(enteredText.Text, originalText.Text))
            {
                digitRepeatLabel.ForeColor = System.Drawing.Color.Green;
                digitRepeatLabel.Text = " Y ";
            }
            else
            {
                digitRepeatLabel.ForeColor = System.Drawing.Color.Red;
                digitRepeatLabel.Text = " N ";
            }
        }

        private void digitOmission_Click(object sender, EventArgs e)
        {
            if (ErrorClassifiers.TestDigitOmission(enteredText.Text, originalText.Text))
            {
                digitOmissionLabel.ForeColor = System.Drawing.Color.Green;
                digitOmissionLabel.Text = " Y ";
            }
            else
            {
                digitOmissionLabel.ForeColor = System.Drawing.Color.Red;
                digitOmissionLabel.Text = " N ";
            }
        }

        private void decimalOmission_Click(object sender, EventArgs e)
        {
            if (ErrorClassifiers.TestDecimalOmission(enteredText.Text, originalText.Text))
            {
                decimalOmissionLabel.ForeColor = System.Drawing.Color.Green;
                decimalOmissionLabel.Text = " Y ";
            }
            else
            {
                decimalOmissionLabel.ForeColor = System.Drawing.Color.Red;
                decimalOmissionLabel.Text = " N ";
            }
        }

        private void extraDigit_Click(object sender, EventArgs e)
        {
            if (ErrorClassifiers.TestExtraDigit(enteredText.Text, originalText.Text))
            {
                extraDigitLabel.ForeColor = System.Drawing.Color.Green;
                extraDigitLabel.Text = " Y ";
            }
            else
            {
                extraDigitLabel.ForeColor = System.Drawing.Color.Red;
                extraDigitLabel.Text = " N ";
            }
        }


        private void wrongDigit_Click(object sender, EventArgs e)
        {
            if (ErrorClassifiers.TestWrongDigit(enteredText.Text, originalText.Text))
            {
                wrongDigitLabel.ForeColor = System.Drawing.Color.Green;
                wrongDigitLabel.Text = " Y ";
            }
            else
            {
                wrongDigitLabel.ForeColor = System.Drawing.Color.Red;
                wrongDigitLabel.Text = " N ";
            }
        }

        private void digitTransposition_Click(object sender, EventArgs e)
        {
            if (ErrorClassifiers.TestDigitTransposition(enteredText.Text, originalText.Text))
            {
                digitTranspositionLabel.ForeColor = System.Drawing.Color.Green;
                digitTranspositionLabel.Text = " Y ";
            }
            else
            {
                digitTranspositionLabel.ForeColor = System.Drawing.Color.Red;
                digitTranspositionLabel.Text = " N ";
            }
        }

        private void signError_Click(object sender, EventArgs e)
        {
            if (ErrorClassifiers.TestSignError(enteredText.Text, originalText.Text))
            {
                signErrorLabel.ForeColor = System.Drawing.Color.Green;
                signErrorLabel.Text = " Y ";
            }
            else
            {
                signErrorLabel.ForeColor = System.Drawing.Color.Red;
                signErrorLabel.Text = " N ";
            }
        }

        private void originalText_TextChanged(object sender, EventArgs e)
        {
            signError_Click(sender, e);
            digitTransposition_Click(sender, e);
            wrongDigit_Click(sender, e);
            extraDigit_Click(sender, e);
            digitOmission_Click(sender, e);
            decimalOmission_Click(sender, e);
            digitRepeat_Click(sender, e);
            decimalPoint_Click(sender, e);
            signOmission_Click(sender, e);
        } //End originalText_TextChanged

        private void enteredText_TextChanged(object sender, EventArgs e)
        {
            signError_Click(sender, e);
            digitTransposition_Click(sender, e);
            wrongDigit_Click(sender, e);
            extraDigit_Click(sender, e);
            digitOmission_Click(sender, e);
            decimalOmission_Click(sender, e);
            digitRepeat_Click(sender, e);
            decimalPoint_Click(sender, e);
            signOmission_Click(sender, e);
        } //End enteredText_TextChanged
    }
}
