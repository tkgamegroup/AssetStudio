using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using AssetStudio;

namespace AssetStudioGUI
{
    public partial class CLIDialog : Form
    {
        int curr_step = 0;
        const int steps = 5;
        public CLIDialog()
        {
            InitializeComponent();

            Progress.Default = new GUIProgress(SetProgressValue, SetTaskName);
        }
        private void SetTaskName(string task)
        {
            curr_step++;
            if (InvokeRequired)
            {
                BeginInvoke(new Action(() => { this.Text = string.Format("Step: {0} ({1}/{2})", task, curr_step, steps); }));
            }
            else
            {
                this.Text = string.Format("Step: {0} ({1}/{2})", task, curr_step, steps);
            }
        }

        private void SetProgressValue(int current, int total)
        {
            var value = (int)(current * 100f / total);
            if (InvokeRequired)
            {
                BeginInvoke(new Action(() => {
                    label1.Text = string.Format("{0}/{1}", current, total);
                    progressBar1.Value = value; 
                }));
            }
            else
            {
                label1.Text = string.Format("{0}/{1}", current, total);
                progressBar1.Value = value;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Progress.stopTask = true;
        }
    }
}
