using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;

namespace ErrorClassifier
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void buttonOpenDialog_Click(object sender, RibbonControlEventArgs e)
        {
            Form1 form = new Form1();
            form.ShowDialog();
        }
    }
}
