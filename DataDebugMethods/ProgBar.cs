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
    /// <summary>
    /// This progress bar's lifecycle should be managed by the UI layer.
    /// </summary>
    public partial class ProgBar : Form
    {
        int Maximum;
        int Minimum;
        int state;

        public ProgBar(int low, int high)
        {
            Minimum = low;
            Maximum = high;
            state = Minimum;
            InitializeComponent();
            progressBar1.Minimum = Minimum;
            progressBar1.Maximum = Maximum;
            this.Visible = true;
        }

        private void ProgBar_Load(object sender, EventArgs e)
        {

        }

        public void SetProgress(int progress)
        {
            if (progress < Minimum || progress > Maximum)
            {
                throw new Exception("Progress bar error.");
            }
            progressBar1.Value = progress;
        }

        public void IncrementProgress(int delta)
        {
            if (state + delta < Minimum)
            {
                state = Minimum;
            }
            else if (state + delta > Maximum)
            {
                state = Maximum;
            }
            else
            {
                state += delta;
            }
            progressBar1.Value = state;
        }

        public int maxProgress()
        {
            return progressBar1.Maximum;
        }
    }
}
