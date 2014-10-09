using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace DataDebugMethods
{
    public class ProgressMaxUnsetException : Exception { }

    /// <summary>
    /// This progress bar's lifecycle should be managed by the UI layer.
    /// </summary>
    public partial class ProgBar : Form
    {
        private bool _max_set = false;
        private int _count = 0;

        public ProgBar()
        {
            InitializeComponent();
            progressBar1.Minimum = 0;
            progressBar1.Maximum = 100;
            this.Visible = true;
        }

        private void ProgBar_Load(object sender, EventArgs e)
        {

        }

        public void IncrementProgress()
        {
            if (!_max_set)
            {
                throw new ProgressMaxUnsetException();
            }

            var iota = 1.0 / progressBar1.Maximum;

            if (_count * iota < 0)
            {
                progressBar1.Value = 0;
            }
            else if (_count * iota > progressBar1.Maximum)
            {
                progressBar1.Value = progressBar1.Maximum;
            }
            else
            {
                progressBar1.Value = (int)(_count * iota);
            }
            _count++;
        }

        public int maxProgress()
        {
            if (!_max_set)
            {
                throw new ProgressMaxUnsetException();
            }
            return progressBar1.Maximum;
        }

        public void setMax(int max_updates)
        {
            progressBar1.Maximum = max_updates;
            _max_set = true;
        }
    }
}
