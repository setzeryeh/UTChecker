using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace UTChecker
{
    public partial class LoggerForm : Form
    {
        private TDS_Parser gTDSParser = null;

        private BackgroundWorker g_bwLogger;
        private BackgroundWorker g_bwProgress;


        public LoggerForm(ref TDS_Parser tdsParser)
        {
            InitializeComponent();

            gTDSParser = tdsParser;


            InitializeBackgroundWorkerForLogger();
            tdsParser.ReportMessageEvent += new ReportMessageEventHandler(this.ReportMessage);
            tdsParser.ClearMessageEvent += new ReportMessageEventHandler(this.ClearMessage);


            InitializeBackgroundWorkerForProgress();
            tdsParser.ReportProgressEvent += new ReportProgressEventHandler(this.ReportProgress);
            tdsParser.ClearProgressEvent += new ReportProgressEventHandler(this.ClearProgress);
        }



        /// <summary>
        /// Init a backgroundworker for log message to listbox
        /// </summary>
        public void InitializeBackgroundWorkerForLogger()
        {
            g_bwLogger = new BackgroundWorker();

            g_bwLogger.WorkerReportsProgress = true;
            g_bwLogger.WorkerSupportsCancellation = true;
            g_bwLogger.DoWork += new DoWorkEventHandler(bwLogger_DoWork);
            g_bwLogger.ProgressChanged += new ProgressChangedEventHandler(bwLogger_ProgressChanged);
            g_bwLogger.RunWorkerCompleted += new RunWorkerCompletedEventHandler(bwLogger_RunWorkerCompleted);

            // run
            g_bwLogger.RunWorkerAsync();
        }

        // event for DoWork
        private void bwLogger_DoWork(object sender, DoWorkEventArgs e)
        {
            while (true)
            {
                if (g_bwLogger.CancellationPending == true)
                {
                    e.Cancel = true;
                    break;
                }
            }
        }

        // event for ProgressChanged
        private void bwLogger_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {

            if (e.ProgressPercentage == 0 && ((string)e.UserState).Equals("ClearMessage"))
            {
                listBoxLogger.Items.Clear();
            }
            else
            {
                string msg = (string)e.UserState;
                this.listBoxLogger.Items.Add(msg);

                int visibleItems = listBoxLogger.ClientSize.Height / listBoxLogger.ItemHeight;
                listBoxLogger.TopIndex = Math.Max(listBoxLogger.Items.Count - visibleItems + 1, 0);
            }

        }

        // event for 
        private void bwLogger_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {

        }


        /// <summary>
        /// ReportMessageEventArgs
        /// </summary>
        public class ReportMessageEventArgs : EventArgs
        {
            public string message { get; set; }
        }

        public delegate void ReportMessageEventHandler(object sender, ReportMessageEventArgs eventArgs);
        public void ReportMessage(object sender, ReportMessageEventArgs e)
        {
            g_bwLogger.ReportProgress(1, e.message);
        }

        public void ClearMessage(object sender, ReportMessageEventArgs e)
        {
            g_bwLogger.ReportProgress(0, "ClearMessage");
        }






        /// <summary>
        /// Init a backgroundworker for progress bar
        /// </summary>
        public void InitializeBackgroundWorkerForProgress()
        {
            g_bwProgress = new BackgroundWorker();

            g_bwProgress.WorkerReportsProgress = true;
            g_bwProgress.WorkerSupportsCancellation = true;
            g_bwProgress.DoWork += new DoWorkEventHandler(bwProgress_DoWork);
            g_bwProgress.ProgressChanged += new ProgressChangedEventHandler(bwProgress_ProgressChanged);
            g_bwProgress.RunWorkerCompleted += new RunWorkerCompletedEventHandler(bwProgress_RunWorkerCompleted);

            // run
            g_bwProgress.RunWorkerAsync();
        }

        // event for DoWork
        private void bwProgress_DoWork(object sender, DoWorkEventArgs e)
        {
            while (true)
            {
                if (g_bwProgress.CancellationPending == true)
                {
                    e.Cancel = true;
                    break;
                }
            }
            
        }

        // event for ProgressChanged
        private void bwProgress_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            if (e.ProgressPercentage == 0 && ((string)e.UserState).Equals("ClearMessage"))
            {
                this.progressBar1.Value = 0;
            }
            else
            {
                this.progressBar1.Value = e.ProgressPercentage;
            }

        }

        // event for 
        private void bwProgress_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {

        }


        /// <summary>
        /// ReportProgressEventArgs
        /// </summary>
        public class ReportProgressEventArgs : EventArgs
        {
            public int progress { get; set; }
        }

        public delegate void ReportProgressEventHandler(object sender, ReportProgressEventArgs eventArgs);
        public void ReportProgress(object sender, ReportProgressEventArgs e)
        {

            g_bwProgress.ReportProgress(e.progress, "");

        }

        public void ClearProgress(object sender, ReportProgressEventArgs e)
        {
            g_bwProgress.ReportProgress(0, "ClearProgress");

        }

        private void LoggerForm_FormClosing(object sender, FormClosingEventArgs e)
        {

            g_bwLogger.CancelAsync();
            g_bwLogger.Dispose();

            g_bwProgress.CancelAsync();
            g_bwProgress.Dispose();

        }

        private void LoggerForm_FormClosed(object sender, FormClosedEventArgs e)
        {
            gTDSParser.ReportMessageEvent -= new ReportMessageEventHandler(this.ReportMessage);
            gTDSParser.ClearMessageEvent -= new ReportMessageEventHandler(this.ClearMessage);

            gTDSParser.ReportProgressEvent -= new ReportProgressEventHandler(this.ReportProgress);
            gTDSParser.ClearProgressEvent -= new ReportProgressEventHandler(this.ClearProgress);


        }
    }
}
