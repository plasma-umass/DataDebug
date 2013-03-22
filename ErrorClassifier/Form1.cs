using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using DataDebugMethods;

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
                MessageBox.Show("Sign omission: YES");
            }
            else
            {
                MessageBox.Show("Sign omission: NO");
            }   
        }

        private void decimalPoint_Click(object sender, EventArgs e)
        {
            if (ErrorClassifiers.TestMisplacedDecimal(enteredText.Text, originalText.Text))
            {
                MessageBox.Show("Misplaced decimal point: YES");
            }
            else
            {
                MessageBox.Show("Misplaced decimal point: NO");
            }   
        }

        private void digitRepeat_Click(object sender, EventArgs e)
        {
            if (ErrorClassifiers.TestDigitRepeat(enteredText.Text, originalText.Text))
            {
                MessageBox.Show("Repeated digit: YES");
            }
            else
            {
                MessageBox.Show("Repeated digit: NO");
            }
        }

        private void digitOmission_Click(object sender, EventArgs e)
        {
            if (ErrorClassifiers.TestDigitOmission(enteredText.Text, originalText.Text))
            {
                MessageBox.Show("Digit omission: YES");
            }
            else
            {
                MessageBox.Show("Digit omission: NO");
            }
        }

        private void decimalOmission_Click(object sender, EventArgs e)
        {
            if (ErrorClassifiers.TestDecimalOmission(enteredText.Text, originalText.Text))
            {
                MessageBox.Show("Decimal omission: YES");
            }
            else
            {
                MessageBox.Show("Decimal omission: NO");
            }
        }

        private void extraDigit_Click(object sender, EventArgs e)
        {
            if (ErrorClassifiers.TestExtraDigit(enteredText.Text, originalText.Text))
            {
                MessageBox.Show("Extra digit: YES");
            }
            else
            {
                MessageBox.Show("Extra digit: NO");
            }
        }


        private void wrongDigit_Click(object sender, EventArgs e)
        {
            if (ErrorClassifiers.TestWrongDigit(enteredText.Text, originalText.Text))
            {
                MessageBox.Show("Wrong digit: YES");
            }
            else
            {
                MessageBox.Show("Wrong digit: NO");
            }
        }
    }
}
