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
                signOmissionTextBox.BackColor = System.Drawing.Color.LightGreen;
                signOmissionTextBox.Text = "Y";
            }
            else
            {
                signOmissionTextBox.BackColor = System.Drawing.Color.Red;
                signOmissionTextBox.Text = "N";
            }
        }

        private void decimalPoint_Click(object sender, EventArgs e)
        {
            if (ErrorClassifiers.TestMisplacedDecimal(enteredText.Text, originalText.Text))
            {
                misplacedDecimalTextBox.BackColor = System.Drawing.Color.LightGreen;
                misplacedDecimalTextBox.Text = "Y";
            }
            else
            {
                misplacedDecimalTextBox.BackColor = System.Drawing.Color.Red;
                misplacedDecimalTextBox.Text = "N";
            }   
        }

        private void digitRepeat_Click(object sender, EventArgs e)
        {
            if (ErrorClassifiers.TestDigitRepeat(enteredText.Text, originalText.Text))
            {
                digitRepeatTextBox.BackColor = System.Drawing.Color.LightGreen;
                digitRepeatTextBox.Text = "Y";
            }
            else
            {
                digitRepeatTextBox.BackColor = System.Drawing.Color.Red;
                digitRepeatTextBox.Text = "N";
            }
        }

        private void digitOmission_Click(object sender, EventArgs e)
        {
            if (ErrorClassifiers.TestDigitOmission(enteredText.Text, originalText.Text))
            {
                digitOmissionTextBox.BackColor = System.Drawing.Color.LightGreen;
                digitOmissionTextBox.Text = "Y";
            }
            else
            {
                digitOmissionTextBox.BackColor = System.Drawing.Color.Red;
                digitOmissionTextBox.Text = "N";
            }
        }

        private void decimalOmission_Click(object sender, EventArgs e)
        {
            if (ErrorClassifiers.TestDecimalOmission(enteredText.Text, originalText.Text))
            {
                decimalOmissionTextBox.BackColor = System.Drawing.Color.LightGreen;
                decimalOmissionTextBox.Text = "Y";
            }
            else
            {
                decimalOmissionTextBox.BackColor = System.Drawing.Color.Red;
                decimalOmissionTextBox.Text = "N";
            }
        }

        private void extraDigit_Click(object sender, EventArgs e)
        {
            if (ErrorClassifiers.TestExtraDigit(enteredText.Text, originalText.Text))
            {
                extraDigitTextBox.BackColor = System.Drawing.Color.LightGreen;
                extraDigitTextBox.Text = "Y";
            }
            else
            {
                extraDigitTextBox.BackColor = System.Drawing.Color.Red;
                extraDigitTextBox.Text = "N";
            }
        }


        private void wrongDigit_Click(object sender, EventArgs e)
        {
            if (ErrorClassifiers.TestWrongDigit(enteredText.Text, originalText.Text))
            {
                wrongDigitTextBox.BackColor = System.Drawing.Color.LightGreen;
                wrongDigitTextBox.Text = "Y";
            }
            else
            {
                wrongDigitTextBox.BackColor = System.Drawing.Color.Red;
                wrongDigitTextBox.Text = "N";
            }
        }

        private void digitTransposition_Click(object sender, EventArgs e)
        {
            if (ErrorClassifiers.TestDigitTransposition(enteredText.Text, originalText.Text))
            {
                digitTranspositionTextBox.BackColor = System.Drawing.Color.LightGreen;
                digitTranspositionTextBox.Text = "Y";
            }
            else
            {
                digitTranspositionTextBox.BackColor = System.Drawing.Color.Red;
                digitTranspositionTextBox.Text = "N";
            }
        }

        private void signError_Click(object sender, EventArgs e)
        {
            if (ErrorClassifiers.TestSignError(enteredText.Text, originalText.Text))
            {
                signErrorTextBox.BackColor = System.Drawing.Color.LightGreen;
                signErrorTextBox.Text = "Y";
            }
            else
            {
                signErrorTextBox.BackColor = System.Drawing.Color.Red;
                signErrorTextBox.Text = "N";
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
