using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;


namespace UTChecker
{
    public partial class MainForm : Form
    {

        UTChecker gUTChecker = null;

        LoggerForm gLoggerForm = null;


        /// <summary>
        /// MainForm body
        /// </summary>
        public MainForm()
        {

            InitializeComponent();

            this.textBoxReferenceListsPath.Hide();
            this.textBoxSUTRRPath.Hide();

            this.buttonSelectReferenceListsPath.Hide();
            this.buttonSelectSUTRRPath.Hide();


            // init a form for Logger/Progress
            gLoggerForm = new LoggerForm();

            // init UT checker
            gUTChecker = new UTChecker();
            gUTChecker.UpdatePathEvent += new EventHandler<UTCheckerEvent>(this.UpdatePath_event);
            gUTChecker.CompletedEvent += new EventHandler<UTCheckerEvent>(this.Completed_event);



        }




        /// <summary>
        /// A event at the MainForm is loaded.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void MainForm_Load(object sender, EventArgs e)
        {
            // get args from command line
            char[] delimiter = new char[] { ' ', '"' };
            string[] args = Environment.CommandLine.Split(delimiter, StringSplitOptions.RemoveEmptyEntries);

            if (!gUTChecker.UpdateEnvironmentSetting(args))
            {
                Environment.ExitCode = -2;
                Environment.Exit(Environment.ExitCode);
            }

            if (gUTChecker.Mode == RunMode.CommandLine)
            {
                if (gLoggerForm.IsDisposed)
                {
                    gLoggerForm = new LoggerForm();
                }

                //gLoggerForm.Focus();
                gLoggerForm.Show();
                gLoggerForm.WindowState = FormWindowState.Minimized;


                // minimized the MainForm
                this.WindowState = FormWindowState.Minimized;
                this.buttonRun.Enabled = false;

                // trigger TDS_Parser
                gUTChecker.Run();
            }

        }


        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void buttonRun_Click(object sender, EventArgs e)
        {
            if (gLoggerForm.IsDisposed)
            {
                gLoggerForm = new LoggerForm();
            }

            gLoggerForm.Show();
            //gLoggerForm.Focus();

            // minimized the MainForm
            this.WindowState = FormWindowState.Minimized;

            // disable the Run button
            this.buttonRun.Enabled = false;


            //triiger TDS_Parser

            gUTChecker.UpdateEnvironmentSetting(GetPath());
            gUTChecker.Run();

        }


        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void buttonSelectListFile_Click(object sender, EventArgs e)
        {
            this.openFileDialog1.Filter = "List Files|*.txt";
            this.openFileDialog1.Title = "Select a List File";

            if (this.openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                this.textBoxListFilePath.Text = this.openFileDialog1.FileName;
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void buttonSelectTDSPath_Click(object sender, EventArgs e)
        {
            if (this.folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                this.textBoxTDSPath.Text = folderBrowserDialog1.SelectedPath;
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void buttonSelectReportTempate_Click(object sender, EventArgs e)
        {
            this.openFileDialog1.Filter = "Excel Files|*.xlsx";
            this.openFileDialog1.Title = "Select a Excel File";

            if (this.openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                this.textBoxReportTemplate.Text = this.openFileDialog1.FileName;
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void buttonSelectSummaryTemplate_Click(object sender, EventArgs e)
        {
            this.openFileDialog1.Filter = "Excel Files|*.xlsx";
            this.openFileDialog1.Title = "Select a Excel File";

            if (this.openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                this.textBoxSummaryTemplate.Text = this.openFileDialog1.FileName;
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void buttonSelectTestLogsPath_Click(object sender, EventArgs e)
        {
            if (this.folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                this.textBoxTestLogPath.Text = this.folderBrowserDialog1.SelectedPath;
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void buttonSelectOutputPath_Click(object sender, EventArgs e)
        {
            if (this.folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                this.textBoxOutputPath.Text = this.folderBrowserDialog1.SelectedPath;
            }
        }


        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void buttonSelectSourceListPath_Click(object sender, EventArgs e)
        {
            if (this.folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                this.textBoxReferenceListsPath.Text = this.folderBrowserDialog1.SelectedPath;
            }
        }


        /// <summary>
        /// folder browser for SUTS
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void buttonSelectSUTSPath_Click(object sender, EventArgs e)
        {
            if (this.folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                this.textBoxSUTSPath.Text = this.folderBrowserDialog1.SelectedPath;
            }
        }


        /// <summary>
        /// folder browser for SUTRR
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void buttonSelectSUTRRPath_Click(object sender, EventArgs e)
        {
            if (this.folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                this.textBoxSUTRRPath.Text = this.folderBrowserDialog1.SelectedPath;
            }
        }


        

        /// <summary>
        /// A method for update the path (delegate method)
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        public void UpdatePath_event(object sender, UTCheckerEvent e)
        {
            // update setting to textbox
            this.textBoxListFilePath.Text = e.Path.listFile;
            this.textBoxTDSPath.Text = e.Path.tdsPath;
            this.textBoxOutputPath.Text = e.Path.outputPath;
            this.textBoxReportTemplate.Text = e.Path.reportTemplate;
            this.textBoxSummaryTemplate.Text = e.Path.summaryTemplate;
            this.textBoxTestLogPath.Text = e.Path.testlogPath;
            this.textBoxReferenceListsPath.Text = e.Path.referenceListsPath;
            this.textBoxSUTSPath.Text = e.Path.sutsPath;
            this.textBoxSUTRRPath.Text = e.Path.sutrrPath;
        }


        /// <summary>
        /// A method for the event of Completed of UTChecker
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        public void Completed_event(object sender, UTCheckerEvent e)
        {

            // enable "Run" Button
            this.buttonRun.Enabled = true;

            // get return code
            int code = e.ReturnCode;

            if (e.Mode == RunMode.CommandLine)
            {
                Environment.ExitCode = code;
                Environment.Exit(Environment.ExitCode);
            }
            else
            {
                string msg = (code == UTChecker.RETURN_CODE.ERROR_USER) ? "Error" : "Done";
                MessageBox.Show(msg + " " + e.Message);

                this.WindowState = FormWindowState.Normal;
                this.Focus();
            }
        }


        /// <summary>
        /// Get the path settings from each textbox.
        /// </summary>
        /// <returns></returns>
        public EnvrionmentSetting GetPath()
        {
            EnvrionmentSetting ps = new EnvrionmentSetting();

            ps.listFile = this.textBoxListFilePath.Text;
            ps.tdsPath = this.textBoxTDSPath.Text;
            ps.outputPath = this.textBoxOutputPath.Text;
            ps.reportTemplate = this.textBoxReportTemplate.Text;
            ps.summaryTemplate = this.textBoxSummaryTemplate.Text;
            ps.testlogPath = this.textBoxTestLogPath.Text;
            ps.referenceListsPath = this.textBoxReferenceListsPath.Text;
            ps.sutsPath = this.textBoxSUTSPath.Text;
            ps.sutrrPath = this.textBoxSUTRRPath.Text;

            return ps;
        }






    }
}
