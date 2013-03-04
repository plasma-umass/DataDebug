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

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            //HERE WE WILL CALL THE FUNCTIONS TO DO THE ANALYSIS
            // 1. CONSTRUCT TREE ==> report progress
            // 2. PERTURBATIONS ==> report progress
            // 3. IMPACT SCORING / LOOK FOR OUTLIERS ==> report progress
            for (int i = this.progressBar1.Minimum; i <= this.progressBar1.Maximum; i++)
            {
                Thread.Sleep(50);
                backgroundWorker1.ReportProgress(i);
            }
        }

        private void backgroundWorker1_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            progressBar1.Value = e.ProgressPercentage;
        }

        
    }
}
