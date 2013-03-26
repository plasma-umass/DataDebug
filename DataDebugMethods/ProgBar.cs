using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Threading;

namespace DataDebugMethods
{
    public partial class ProgBar : Form
    {
        public ProgBar(int min, int max)
        {
            this.Visible = true;
            InitializeComponent();
            progressBar1.Minimum = min;
            progressBar1.Maximum = max;
            progressBar1.Value = progressBar1.Minimum;
        }

        private void ProgBar_Load(object sender, System.EventArgs e)
        {
        }

        public void SetProgress(int progress)
        {
            progressBar1.Value = progress;
        }

        public int maxProgress()
        {
            return progressBar1.Maximum;
        }
    }
}
