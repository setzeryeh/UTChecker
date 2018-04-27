namespace UTChecker
{
    partial class MainForm
    {
        /// <summary>
        /// 設計工具所需的變數。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清除任何使用中的資源。
        /// </summary>
        /// <param name="disposing">如果應該處置受控資源則為 true，否則為 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form 設計工具產生的程式碼

        /// <summary>
        /// 此為設計工具支援所需的方法 - 請勿使用程式碼編輯器修改
        /// 這個方法的內容。
        /// </summary>
        private void InitializeComponent()
        {
            this.textBoxListFilePath = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.textBoxTDSPath = new System.Windows.Forms.TextBox();
            this.textBoxReportTemplate = new System.Windows.Forms.TextBox();
            this.textBoxSummaryTemplate = new System.Windows.Forms.TextBox();
            this.textBoxOutputPath = new System.Windows.Forms.TextBox();
            this.groupBoxEnvironment = new System.Windows.Forms.GroupBox();
            this.buttonSelectTestLogsPath = new System.Windows.Forms.Button();
            this.buttonSelectSummaryTemplate = new System.Windows.Forms.Button();
            this.buttonSelectReportTempate = new System.Windows.Forms.Button();
            this.buttonSelectListFile = new System.Windows.Forms.Button();
            this.buttonSelectTDSPath = new System.Windows.Forms.Button();
            this.buttonSelectOutputPath = new System.Windows.Forms.Button();
            this.label6 = new System.Windows.Forms.Label();
            this.textBoxTestLogs = new System.Windows.Forms.TextBox();
            this.buttonRun = new System.Windows.Forms.Button();
            this.folderBrowserDialog1 = new System.Windows.Forms.FolderBrowserDialog();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.groupBoxEnvironment.SuspendLayout();
            this.SuspendLayout();
            // 
            // textBoxListFilePath
            // 
            this.textBoxListFilePath.Location = new System.Drawing.Point(134, 25);
            this.textBoxListFilePath.Name = "textBoxListFilePath";
            this.textBoxListFilePath.Size = new System.Drawing.Size(424, 22);
            this.textBoxListFilePath.TabIndex = 0;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(10, 28);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(53, 14);
            this.label1.TabIndex = 1;
            this.label1.Text = "List File:";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(10, 63);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(57, 14);
            this.label2.TabIndex = 1;
            this.label2.Text = "TDS Path:";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(10, 98);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(93, 14);
            this.label3.TabIndex = 1;
            this.label3.Text = "Report Templte:";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(10, 133);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(113, 14);
            this.label4.TabIndex = 1;
            this.label4.Text = "Summary Template:";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(10, 214);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(74, 14);
            this.label5.TabIndex = 1;
            this.label5.Text = "Output Path:";
            // 
            // textBoxTDSPath
            // 
            this.textBoxTDSPath.Location = new System.Drawing.Point(134, 60);
            this.textBoxTDSPath.Name = "textBoxTDSPath";
            this.textBoxTDSPath.Size = new System.Drawing.Size(424, 22);
            this.textBoxTDSPath.TabIndex = 0;
            // 
            // textBoxReportTemplate
            // 
            this.textBoxReportTemplate.Location = new System.Drawing.Point(134, 95);
            this.textBoxReportTemplate.Name = "textBoxReportTemplate";
            this.textBoxReportTemplate.Size = new System.Drawing.Size(424, 22);
            this.textBoxReportTemplate.TabIndex = 0;
            // 
            // textBoxSummaryTemplate
            // 
            this.textBoxSummaryTemplate.Location = new System.Drawing.Point(135, 130);
            this.textBoxSummaryTemplate.Name = "textBoxSummaryTemplate";
            this.textBoxSummaryTemplate.Size = new System.Drawing.Size(424, 22);
            this.textBoxSummaryTemplate.TabIndex = 0;
            // 
            // textBoxOutputPath
            // 
            this.textBoxOutputPath.Location = new System.Drawing.Point(135, 211);
            this.textBoxOutputPath.Name = "textBoxOutputPath";
            this.textBoxOutputPath.Size = new System.Drawing.Size(424, 22);
            this.textBoxOutputPath.TabIndex = 0;
            // 
            // groupBoxEnvironment
            // 
            this.groupBoxEnvironment.BackColor = System.Drawing.SystemColors.Control;
            this.groupBoxEnvironment.Controls.Add(this.buttonSelectTestLogsPath);
            this.groupBoxEnvironment.Controls.Add(this.buttonSelectSummaryTemplate);
            this.groupBoxEnvironment.Controls.Add(this.buttonSelectReportTempate);
            this.groupBoxEnvironment.Controls.Add(this.buttonSelectListFile);
            this.groupBoxEnvironment.Controls.Add(this.buttonSelectTDSPath);
            this.groupBoxEnvironment.Controls.Add(this.buttonSelectOutputPath);
            this.groupBoxEnvironment.Controls.Add(this.label6);
            this.groupBoxEnvironment.Controls.Add(this.textBoxTestLogs);
            this.groupBoxEnvironment.Controls.Add(this.textBoxListFilePath);
            this.groupBoxEnvironment.Controls.Add(this.textBoxTDSPath);
            this.groupBoxEnvironment.Controls.Add(this.label5);
            this.groupBoxEnvironment.Controls.Add(this.textBoxReportTemplate);
            this.groupBoxEnvironment.Controls.Add(this.label4);
            this.groupBoxEnvironment.Controls.Add(this.textBoxSummaryTemplate);
            this.groupBoxEnvironment.Controls.Add(this.label3);
            this.groupBoxEnvironment.Controls.Add(this.textBoxOutputPath);
            this.groupBoxEnvironment.Controls.Add(this.label2);
            this.groupBoxEnvironment.Controls.Add(this.label1);
            this.groupBoxEnvironment.Location = new System.Drawing.Point(12, 12);
            this.groupBoxEnvironment.Name = "groupBoxEnvironment";
            this.groupBoxEnvironment.Size = new System.Drawing.Size(614, 281);
            this.groupBoxEnvironment.TabIndex = 3;
            this.groupBoxEnvironment.TabStop = false;
            this.groupBoxEnvironment.Text = "Environment";
            // 
            // buttonSelectTestLogsPath
            // 
            this.buttonSelectTestLogsPath.Location = new System.Drawing.Point(565, 165);
            this.buttonSelectTestLogsPath.Name = "buttonSelectTestLogsPath";
            this.buttonSelectTestLogsPath.Size = new System.Drawing.Size(24, 22);
            this.buttonSelectTestLogsPath.TabIndex = 4;
            this.buttonSelectTestLogsPath.Text = "...";
            this.buttonSelectTestLogsPath.UseVisualStyleBackColor = true;
            this.buttonSelectTestLogsPath.Click += new System.EventHandler(this.buttonSelectTestLogsPath_Click);
            // 
            // buttonSelectSummaryTemplate
            // 
            this.buttonSelectSummaryTemplate.Location = new System.Drawing.Point(565, 130);
            this.buttonSelectSummaryTemplate.Name = "buttonSelectSummaryTemplate";
            this.buttonSelectSummaryTemplate.Size = new System.Drawing.Size(24, 22);
            this.buttonSelectSummaryTemplate.TabIndex = 4;
            this.buttonSelectSummaryTemplate.Text = "...";
            this.buttonSelectSummaryTemplate.UseVisualStyleBackColor = true;
            this.buttonSelectSummaryTemplate.Click += new System.EventHandler(this.buttonSelectSummaryTemplate_Click);
            // 
            // buttonSelectReportTempate
            // 
            this.buttonSelectReportTempate.Location = new System.Drawing.Point(565, 95);
            this.buttonSelectReportTempate.Name = "buttonSelectReportTempate";
            this.buttonSelectReportTempate.Size = new System.Drawing.Size(24, 22);
            this.buttonSelectReportTempate.TabIndex = 4;
            this.buttonSelectReportTempate.Text = "...";
            this.buttonSelectReportTempate.UseVisualStyleBackColor = true;
            this.buttonSelectReportTempate.Click += new System.EventHandler(this.buttonSelectReportTempate_Click);
            // 
            // buttonSelectListFile
            // 
            this.buttonSelectListFile.Location = new System.Drawing.Point(565, 25);
            this.buttonSelectListFile.Name = "buttonSelectListFile";
            this.buttonSelectListFile.Size = new System.Drawing.Size(24, 22);
            this.buttonSelectListFile.TabIndex = 4;
            this.buttonSelectListFile.Text = "...";
            this.buttonSelectListFile.UseVisualStyleBackColor = true;
            this.buttonSelectListFile.Click += new System.EventHandler(this.buttonSelectListFile_Click);
            // 
            // buttonSelectTDSPath
            // 
            this.buttonSelectTDSPath.Location = new System.Drawing.Point(565, 60);
            this.buttonSelectTDSPath.Name = "buttonSelectTDSPath";
            this.buttonSelectTDSPath.Size = new System.Drawing.Size(24, 22);
            this.buttonSelectTDSPath.TabIndex = 4;
            this.buttonSelectTDSPath.Text = "...";
            this.buttonSelectTDSPath.UseVisualStyleBackColor = true;
            this.buttonSelectTDSPath.Click += new System.EventHandler(this.buttonSelectTDSPath_Click);
            // 
            // buttonSelectOutputPath
            // 
            this.buttonSelectOutputPath.Location = new System.Drawing.Point(565, 211);
            this.buttonSelectOutputPath.Name = "buttonSelectOutputPath";
            this.buttonSelectOutputPath.Size = new System.Drawing.Size(24, 22);
            this.buttonSelectOutputPath.TabIndex = 4;
            this.buttonSelectOutputPath.Text = "...";
            this.buttonSelectOutputPath.UseVisualStyleBackColor = true;
            this.buttonSelectOutputPath.Click += new System.EventHandler(this.buttonSelectOutputPath_Click);
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(10, 168);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(86, 14);
            this.label6.TabIndex = 3;
            this.label6.Text = "Test Logs Path:";
            // 
            // textBoxTestLogs
            // 
            this.textBoxTestLogs.Location = new System.Drawing.Point(135, 165);
            this.textBoxTestLogs.Name = "textBoxTestLogs";
            this.textBoxTestLogs.Size = new System.Drawing.Size(424, 22);
            this.textBoxTestLogs.TabIndex = 2;
            // 
            // buttonRun
            // 
            this.buttonRun.Location = new System.Drawing.Point(647, 247);
            this.buttonRun.Name = "buttonRun";
            this.buttonRun.Size = new System.Drawing.Size(98, 46);
            this.buttonRun.TabIndex = 4;
            this.buttonRun.Text = "Run";
            this.buttonRun.UseVisualStyleBackColor = true;
            this.buttonRun.Click += new System.EventHandler(this.buttonRun_Click);
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 14F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(758, 305);
            this.Controls.Add(this.buttonRun);
            this.Controls.Add(this.groupBoxEnvironment);
            this.Font = new System.Drawing.Font("Calibri", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Name = "MainForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "UT Checker";
            this.Load += new System.EventHandler(this.MainForm_Load);
            this.groupBoxEnvironment.ResumeLayout(false);
            this.groupBoxEnvironment.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TextBox textBoxListFilePath;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.TextBox textBoxTDSPath;
        private System.Windows.Forms.TextBox textBoxReportTemplate;
        private System.Windows.Forms.TextBox textBoxSummaryTemplate;
        private System.Windows.Forms.TextBox textBoxOutputPath;
        private System.Windows.Forms.GroupBox groupBoxEnvironment;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.TextBox textBoxTestLogs;
        private System.Windows.Forms.Button buttonRun;
        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog1;
        private System.Windows.Forms.Button buttonSelectTestLogsPath;
        private System.Windows.Forms.Button buttonSelectSummaryTemplate;
        private System.Windows.Forms.Button buttonSelectReportTempate;
        private System.Windows.Forms.Button buttonSelectListFile;
        private System.Windows.Forms.Button buttonSelectTDSPath;
        private System.Windows.Forms.Button buttonSelectOutputPath;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
    }
}

