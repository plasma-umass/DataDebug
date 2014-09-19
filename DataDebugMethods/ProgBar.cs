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

        delegate void SetProgressCallback(int progress);

        public void SetProgress(int progress)
        {
            // InvokeRequired ensures that the thread accessing
            // the progressbar value is the same as the one
            // that created it.  Forms requires this to ensure
            // thread-safety.
            if (this.progressBar1.InvokeRequired)
            {
                var spc = new SetProgressCallback(SetProgress);
                this.Invoke(spc, new object[] { progress });
            }
            else
            {
                if (progress < Minimum || progress > Maximum)
                {
                    throw new Exception("Progress bar error.");
                }
                progressBar1.Value = progress;
            }
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
