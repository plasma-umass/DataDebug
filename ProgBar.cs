using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Threading;

namespace DataDebug
{
    public partial class ProgBar : Form
    {
        public ProgBar(int min, int max)
        {
            this.Visible = true;
            InitializeComponent();
            progressBar1.Minimum = min;
            progressBar1.Maximum = max;
            // Start the BackgroundWorker.
            backgroundWorker1 = new BackgroundWorker();
            backgroundWorker1.DoWork += new DoWorkEventHandler(backgroundWorker1_DoWork);
            backgroundWorker1.ProgressChanged += new ProgressChangedEventHandler(backgroundWorker1_ProgressChanged);
            backgroundWorker1.WorkerReportsProgress = true;
            progressBar1.Value = progressBar1.Minimum;
            backgroundWorker1.RunWorkerAsync();
        }

        private void ProgBar_Load(object sender, System.EventArgs e)
        {
        }

        public void SetProgress(int progress)
        {
            progressBar1.Value = progress;
        }

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            while (progressBar1.Value <= progressBar1.Maximum)
            {
                backgroundWorker1.ReportProgress(progressBar1.Value);
                Thread.Sleep(10);
                if (progressBar1.Value == progressBar1.Maximum)
                {
                    return;
                }
            }
        }

        private void backgroundWorker1_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            progressBar1.Value = e.ProgressPercentage;
        }
    }
}
