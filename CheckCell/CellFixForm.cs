using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace CheckCell
{
    public partial class CellFixForm : Form
    {
        Microsoft.Office.Interop.Excel.Range _cell;
        System.Drawing.Color _color;
        Action _fn;

        public CellFixForm(Microsoft.Office.Interop.Excel.Range cell, System.Drawing.Color color, Action ReAnalyzeFn)
        {
            _cell = cell;
            _color = color;
            _fn = ReAnalyzeFn;
            InitializeComponent();
        }

        private void CancelFix_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void AcceptFix_Click(object sender, EventArgs e)
        {
            // change the cell value
            _cell.Value2 = this.FixText.Text;

            // change color
            _cell.Interior.Color = _color;

            // close form
            this.Close();

            // call callback
            _fn.Invoke();
        }
    }
}
