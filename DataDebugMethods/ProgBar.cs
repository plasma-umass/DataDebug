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
        int _maximum;
        int _minimum;
        int _state;
        int _poke_count = 0;

        public ProgBar(int low, int high)
        {
            _minimum = low;
            _maximum = high;
            _state = _minimum;
            InitializeComponent();
            progressBar1.Minimum = _minimum;
            progressBar1.Maximum = _maximum;
            this.Visible = true;
        }

        private void ProgBar_Load(object sender, EventArgs e)
        {

        }

        delegate void SetProgressCallback(int progress);

        public void SetProgress(int progress)
        {
            if (progress < _minimum || progress > _maximum)
            {
                throw new Exception("Progress bar error.");
            }
            _state = progress;
            progressBar1.Value = _state;
        }

        public void IncrementProgress(int delta)
        {
            if (_state + delta < _minimum)
            {
                _state = _minimum;
            }
            else if (_state + delta > _maximum)
            {
                _state = _maximum;
            }
            else
            {
                _state += delta;
            }
            progressBar1.Value = _state;
        }

        public int maxProgress()
        {
            return progressBar1.Maximum;
        }

        public void setMax(int m)
        {
            _maximum = m;
        }

        public void pokePB()
        {
            _poke_count += 1;
            this.SetProgress(_poke_count * 100 / _maximum);
        }
    }
}
