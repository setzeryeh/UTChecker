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

        TDS_Parser gTDSParser = null;

        LoggerForm gLoggerForm = null;


        /// <summary>
        /// MainForm body
        /// </summary>
        public MainForm()
        {

            InitializeComponent();

            gTDSParser = new TDS_Parser(this, gLoggerForm);
            gTDSParser.UpdatePathEvent += new EventHandler(this.UpdatePath);

            // init a form for Logger/Progress
            gLoggerForm = new LoggerForm(ref gTDSParser);

        }


        /// <summary>
        /// A event at the MainForm is loaded.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void MainForm_Load(object sender, EventArgs e)
        {

            if (!gTDSParser.UpdatePathSetting())
            {
                
                Environment.ExitCode = 111;
                Environment.Exit(Environment.ExitCode);
            }

            if (gTDSParser.RunUTCheckerBy == TDS_Parser.RunBy.CommandLine)
            {
                if (gLoggerForm.IsDisposed)
                {
                    gLoggerForm = new LoggerForm(ref gTDSParser);
                }


                this.WindowState = FormWindowState.Minimized;
                gLoggerForm.Focus();
                gLoggerForm.Show();
                
                // trigger TDS_Parser
                gTDSParser.Run();
            }

        }


        public delegate void UpdatePathEventHandler(TDS_Parser.PathSetting ps);

        /// <summary>
        /// A method for update the path (delegate method)
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        public void UpdatePath(object sender, EventArgs e)
        {
            // update setting to textbox
            this.textBoxListFilePath.Text = gTDSParser.g_FilePathSetting.listFile;
            this.textBoxTDSPath.Text = gTDSParser.g_FilePathSetting.tdsPath;
            this.textBoxOutputPath.Text = gTDSParser.g_FilePathSetting.outputPath;
            this.textBoxReportTemplate.Text = gTDSParser.g_FilePathSetting.reportTemplate;
            this.textBoxSummaryTemplate.Text = gTDSParser.g_FilePathSetting.summaryTemplate;
            this.textBoxTestLogs.Text = gTDSParser.g_FilePathSetting.testlogsPath;

        }

        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public TDS_Parser.PathSetting GetPath()
        {
            TDS_Parser.PathSetting ps = new TDS_Parser.PathSetting();

            ps.listFile = this.textBoxListFilePath.Text;
            ps.tdsPath = this.textBoxTDSPath.Text;
            ps.outputPath = this.textBoxOutputPath.Text;
            ps.reportTemplate = this.textBoxReportTemplate.Text;
            ps.summaryTemplate = this.textBoxSummaryTemplate.Text;
            ps.testlogsPath = this.textBoxTestLogs.Text;

            return ps;
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
                gLoggerForm = new LoggerForm(ref gTDSParser);
            }

            gLoggerForm.Show();
            gLoggerForm.Focus();

            //triiger TDS_Parser
            gTDSParser.Run();

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
                this.textBoxTestLogs.Text = this.folderBrowserDialog1.SelectedPath;
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
    }
}
