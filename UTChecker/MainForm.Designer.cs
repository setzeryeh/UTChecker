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
            this.buttonSelectReferenceListsPath = new System.Windows.Forms.Button();
            this.buttonSelectSUTRRPath = new System.Windows.Forms.Button();
            this.buttonSelectSUTSPath = new System.Windows.Forms.Button();
            this.buttonSelectOutputPath = new System.Windows.Forms.Button();
            this.label6 = new System.Windows.Forms.Label();
            this.textBoxTestLogPath = new System.Windows.Forms.TextBox();
            this.label7 = new System.Windows.Forms.Label();
            this.label9 = new System.Windows.Forms.Label();
            this.label8 = new System.Windows.Forms.Label();
            this.textBoxReferenceListsPath = new System.Windows.Forms.TextBox();
            this.textBoxSUTRRPath = new System.Windows.Forms.TextBox();
            this.textBoxSUTSPath = new System.Windows.Forms.TextBox();
            this.buttonRun = new System.Windows.Forms.Button();
            this.folderBrowserDialog1 = new System.Windows.Forms.FolderBrowserDialog();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.groupBoxEnvironment.SuspendLayout();
            this.SuspendLayout();
            // 
            // textBoxListFilePath
            // 
            this.textBoxListFilePath.Location = new System.Drawing.Point(132, 25);
            this.textBoxListFilePath.Name = "textBoxListFilePath";
            this.textBoxListFilePath.Size = new System.Drawing.Size(463, 22);
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
            this.label5.Location = new System.Drawing.Point(10, 338);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(74, 14);
            this.label5.TabIndex = 1;
            this.label5.Text = "Output Path:";
            // 
            // textBoxTDSPath
            // 
            this.textBoxTDSPath.Location = new System.Drawing.Point(132, 60);
            this.textBoxTDSPath.Name = "textBoxTDSPath";
            this.textBoxTDSPath.Size = new System.Drawing.Size(463, 22);
            this.textBoxTDSPath.TabIndex = 0;
            // 
            // textBoxReportTemplate
            // 
            this.textBoxReportTemplate.Location = new System.Drawing.Point(132, 95);
            this.textBoxReportTemplate.Name = "textBoxReportTemplate";
            this.textBoxReportTemplate.Size = new System.Drawing.Size(463, 22);
            this.textBoxReportTemplate.TabIndex = 0;
            // 
            // textBoxSummaryTemplate
            // 
            this.textBoxSummaryTemplate.Location = new System.Drawing.Point(133, 130);
            this.textBoxSummaryTemplate.Name = "textBoxSummaryTemplate";
            this.textBoxSummaryTemplate.Size = new System.Drawing.Size(463, 22);
            this.textBoxSummaryTemplate.TabIndex = 0;
            // 
            // textBoxOutputPath
            // 
            this.textBoxOutputPath.Location = new System.Drawing.Point(133, 335);
            this.textBoxOutputPath.Name = "textBoxOutputPath";
            this.textBoxOutputPath.Size = new System.Drawing.Size(463, 22);
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
            this.groupBoxEnvironment.Controls.Add(this.buttonSelectReferenceListsPath);
            this.groupBoxEnvironment.Controls.Add(this.buttonSelectSUTRRPath);
            this.groupBoxEnvironment.Controls.Add(this.buttonSelectSUTSPath);
            this.groupBoxEnvironment.Controls.Add(this.buttonSelectOutputPath);
            this.groupBoxEnvironment.Controls.Add(this.label6);
            this.groupBoxEnvironment.Controls.Add(this.textBoxTestLogPath);
            this.groupBoxEnvironment.Controls.Add(this.textBoxListFilePath);
            this.groupBoxEnvironment.Controls.Add(this.textBoxTDSPath);
            this.groupBoxEnvironment.Controls.Add(this.label7);
            this.groupBoxEnvironment.Controls.Add(this.label9);
            this.groupBoxEnvironment.Controls.Add(this.label8);
            this.groupBoxEnvironment.Controls.Add(this.label5);
            this.groupBoxEnvironment.Controls.Add(this.textBoxReportTemplate);
            this.groupBoxEnvironment.Controls.Add(this.label4);
            this.groupBoxEnvironment.Controls.Add(this.textBoxSummaryTemplate);
            this.groupBoxEnvironment.Controls.Add(this.textBoxReferenceListsPath);
            this.groupBoxEnvironment.Controls.Add(this.textBoxSUTRRPath);
            this.groupBoxEnvironment.Controls.Add(this.textBoxSUTSPath);
            this.groupBoxEnvironment.Controls.Add(this.label3);
            this.groupBoxEnvironment.Controls.Add(this.textBoxOutputPath);
            this.groupBoxEnvironment.Controls.Add(this.label2);
            this.groupBoxEnvironment.Controls.Add(this.label1);
            this.groupBoxEnvironment.Location = new System.Drawing.Point(12, 12);
            this.groupBoxEnvironment.Name = "groupBoxEnvironment";
            this.groupBoxEnvironment.Size = new System.Drawing.Size(643, 378);
            this.groupBoxEnvironment.TabIndex = 3;
            this.groupBoxEnvironment.TabStop = false;
            this.groupBoxEnvironment.Text = "Environment Setting";
            // 
            // buttonSelectTestLogsPath
            // 
            this.buttonSelectTestLogsPath.Location = new System.Drawing.Point(602, 165);
            this.buttonSelectTestLogsPath.Name = "buttonSelectTestLogsPath";
            this.buttonSelectTestLogsPath.Size = new System.Drawing.Size(24, 22);
            this.buttonSelectTestLogsPath.TabIndex = 4;
            this.buttonSelectTestLogsPath.Text = "...";
            this.buttonSelectTestLogsPath.UseVisualStyleBackColor = true;
            this.buttonSelectTestLogsPath.Click += new System.EventHandler(this.buttonSelectTestLogsPath_Click);
            // 
            // buttonSelectSummaryTemplate
            // 
            this.buttonSelectSummaryTemplate.Location = new System.Drawing.Point(602, 130);
            this.buttonSelectSummaryTemplate.Name = "buttonSelectSummaryTemplate";
            this.buttonSelectSummaryTemplate.Size = new System.Drawing.Size(24, 22);
            this.buttonSelectSummaryTemplate.TabIndex = 4;
            this.buttonSelectSummaryTemplate.Text = "...";
            this.buttonSelectSummaryTemplate.UseVisualStyleBackColor = true;
            this.buttonSelectSummaryTemplate.Click += new System.EventHandler(this.buttonSelectSummaryTemplate_Click);
            // 
            // buttonSelectReportTempate
            // 
            this.buttonSelectReportTempate.Location = new System.Drawing.Point(602, 95);
            this.buttonSelectReportTempate.Name = "buttonSelectReportTempate";
            this.buttonSelectReportTempate.Size = new System.Drawing.Size(24, 22);
            this.buttonSelectReportTempate.TabIndex = 4;
            this.buttonSelectReportTempate.Text = "...";
            this.buttonSelectReportTempate.UseVisualStyleBackColor = true;
            this.buttonSelectReportTempate.Click += new System.EventHandler(this.buttonSelectReportTempate_Click);
            // 
            // buttonSelectListFile
            // 
            this.buttonSelectListFile.Location = new System.Drawing.Point(602, 25);
            this.buttonSelectListFile.Name = "buttonSelectListFile";
            this.buttonSelectListFile.Size = new System.Drawing.Size(24, 22);
            this.buttonSelectListFile.TabIndex = 4;
            this.buttonSelectListFile.Text = "...";
            this.buttonSelectListFile.UseVisualStyleBackColor = true;
            this.buttonSelectListFile.Click += new System.EventHandler(this.buttonSelectListFile_Click);
            // 
            // buttonSelectTDSPath
            // 
            this.buttonSelectTDSPath.Location = new System.Drawing.Point(602, 60);
            this.buttonSelectTDSPath.Name = "buttonSelectTDSPath";
            this.buttonSelectTDSPath.Size = new System.Drawing.Size(24, 22);
            this.buttonSelectTDSPath.TabIndex = 4;
            this.buttonSelectTDSPath.Text = "...";
            this.buttonSelectTDSPath.UseVisualStyleBackColor = true;
            this.buttonSelectTDSPath.Click += new System.EventHandler(this.buttonSelectTDSPath_Click);
            // 
            // buttonSelectReferenceListsPath
            // 
            this.buttonSelectReferenceListsPath.Location = new System.Drawing.Point(602, 204);
            this.buttonSelectReferenceListsPath.Name = "buttonSelectReferenceListsPath";
            this.buttonSelectReferenceListsPath.Size = new System.Drawing.Size(24, 22);
            this.buttonSelectReferenceListsPath.TabIndex = 4;
            this.buttonSelectReferenceListsPath.Text = "...";
            this.buttonSelectReferenceListsPath.UseVisualStyleBackColor = true;
            this.buttonSelectReferenceListsPath.Click += new System.EventHandler(this.buttonSelectSourceListPath_Click);
            // 
            // buttonSelectSUTRRPath
            // 
            this.buttonSelectSUTRRPath.Location = new System.Drawing.Point(602, 281);
            this.buttonSelectSUTRRPath.Name = "buttonSelectSUTRRPath";
            this.buttonSelectSUTRRPath.Size = new System.Drawing.Size(24, 22);
            this.buttonSelectSUTRRPath.TabIndex = 4;
            this.buttonSelectSUTRRPath.Text = "...";
            this.buttonSelectSUTRRPath.UseVisualStyleBackColor = true;
            this.buttonSelectSUTRRPath.Click += new System.EventHandler(this.buttonSelectSUTRRPath_Click);
            // 
            // buttonSelectSUTSPath
            // 
            this.buttonSelectSUTSPath.Location = new System.Drawing.Point(602, 242);
            this.buttonSelectSUTSPath.Name = "buttonSelectSUTSPath";
            this.buttonSelectSUTSPath.Size = new System.Drawing.Size(24, 22);
            this.buttonSelectSUTSPath.TabIndex = 4;
            this.buttonSelectSUTSPath.Text = "...";
            this.buttonSelectSUTSPath.UseVisualStyleBackColor = true;
            this.buttonSelectSUTSPath.Click += new System.EventHandler(this.buttonSelectSUTSPath_Click);
            // 
            // buttonSelectOutputPath
            // 
            this.buttonSelectOutputPath.Location = new System.Drawing.Point(602, 335);
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
            this.label6.Size = new System.Drawing.Size(80, 14);
            this.label6.TabIndex = 3;
            this.label6.Text = "Test Log Path:";
            // 
            // textBoxTestLogPath
            // 
            this.textBoxTestLogPath.Location = new System.Drawing.Point(133, 165);
            this.textBoxTestLogPath.Name = "textBoxTestLogPath";
            this.textBoxTestLogPath.Size = new System.Drawing.Size(463, 22);
            this.textBoxTestLogPath.TabIndex = 2;
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(10, 207);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(120, 14);
            this.label7.TabIndex = 1;
            this.label7.Text = "Reference Lists Path:";
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(10, 284);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(64, 14);
            this.label9.TabIndex = 1;
            this.label9.Text = "SUTR Path:";
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(11, 245);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(63, 14);
            this.label8.TabIndex = 1;
            this.label8.Text = "SUTS Path:";
            // 
            // textBoxReferenceListsPath
            // 
            this.textBoxReferenceListsPath.Location = new System.Drawing.Point(133, 204);
            this.textBoxReferenceListsPath.Name = "textBoxReferenceListsPath";
            this.textBoxReferenceListsPath.Size = new System.Drawing.Size(463, 22);
            this.textBoxReferenceListsPath.TabIndex = 0;
            // 
            // textBoxSUTRRPath
            // 
            this.textBoxSUTRRPath.Location = new System.Drawing.Point(132, 281);
            this.textBoxSUTRRPath.Name = "textBoxSUTRRPath";
            this.textBoxSUTRRPath.Size = new System.Drawing.Size(463, 22);
            this.textBoxSUTRRPath.TabIndex = 0;
            // 
            // textBoxSUTSPath
            // 
            this.textBoxSUTSPath.Location = new System.Drawing.Point(132, 242);
            this.textBoxSUTSPath.Name = "textBoxSUTSPath";
            this.textBoxSUTSPath.Size = new System.Drawing.Size(463, 22);
            this.textBoxSUTSPath.TabIndex = 0;
            // 
            // buttonRun
            // 
            this.buttonRun.Location = new System.Drawing.Point(746, 344);
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
            this.ClientSize = new System.Drawing.Size(874, 409);
            this.Controls.Add(this.buttonRun);
            this.Controls.Add(this.groupBoxEnvironment);
            this.Font = new System.Drawing.Font("Calibri", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Name = "MainForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "UT Checker 1.00";
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
        private System.Windows.Forms.TextBox textBoxTestLogPath;
        private System.Windows.Forms.Button buttonRun;
        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog1;
        private System.Windows.Forms.Button buttonSelectTestLogsPath;
        private System.Windows.Forms.Button buttonSelectSummaryTemplate;
        private System.Windows.Forms.Button buttonSelectReportTempate;
        private System.Windows.Forms.Button buttonSelectListFile;
        private System.Windows.Forms.Button buttonSelectTDSPath;
        private System.Windows.Forms.Button buttonSelectOutputPath;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.Button buttonSelectReferenceListsPath;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.TextBox textBoxReferenceListsPath;
        private System.Windows.Forms.Button buttonSelectSUTRRPath;
        private System.Windows.Forms.Button buttonSelectSUTSPath;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.TextBox textBoxSUTRRPath;
        private System.Windows.Forms.TextBox textBoxSUTSPath;
    }
}

