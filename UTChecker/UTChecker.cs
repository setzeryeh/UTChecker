using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;

namespace UTChecker
{
    public partial class UTChecker
    {

        /// <summary>
        /// Constructor for UTChecker
        /// </summary>
        public UTChecker()
        {

            // init a task of backgroundworker for UTChecker
            InitializeBackgroundWorkerForUTChecker();


        }



        #region BackgroundWorker for UT Checker

        /// <summary>
        /// Init a backgroundworker for log message to listbox
        /// </summary>
        public void InitializeBackgroundWorkerForUTChecker()
        {
            g_bwUTChecker = new BackgroundWorker();

            //g_bwUTChecker.WorkerReportsProgress = true;
            //g_bwUTChecker.WorkerSupportsCancellation = true;
            g_bwUTChecker.DoWork += new DoWorkEventHandler(bwUTChecker_DoWork);
            //g_bwUTChecker.ProgressChanged += new ProgressChangedEventHandler(bwUTChecker_ProgressChanged);
            g_bwUTChecker.RunWorkerCompleted += new RunWorkerCompletedEventHandler(bwUTChecker_RunWorkerCompleted);
        }



        ///// <summary>
        ///// 
        ///// </summary>
        //public void Stop()
        //{
        //    g_bwUTChecker.CancelAsync();
        //}



        /// <summary>
        /// An event which triggers the RunTDSParser
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void bwUTChecker_DoWork(object sender, DoWorkEventArgs e)
        {

            // Get the BackgroundWorker that raised this event.
            BackgroundWorker worker = sender as BackgroundWorker;

            // call RunUTchecker
            e.Result = RunUTChecker();

        }


        ///// <summary>
        ///// bwUTChecker_ProgressChanged
        ///// </summary>
        ///// <param name="sender"></param>
        ///// <param name="e"></param>
        //private void bwUTChecker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        //{
        //}


        /// <summary>
        /// bwUTChecker_RunWorkerCompleted
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void bwUTChecker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            
            if (e.Error != null)
            {
                MessageBox.Show(e.Error.Message);
            }
            else if (e.Cancelled)
            {
                MessageBox.Show("No support Cancel!");
            }

            int result = (int)e.Result;

            if (result == -1)
            {
                if (this.Mode == RunMode.CommandLine)
                {
                    this.ReturnCode = RETURN_CODE.ERROR_CMD;
                }
                else
                {
                    this.ReturnCode = RETURN_CODE.ERROR_USER;
                }
            }
            else
            {
                this.ReturnCode = RETURN_CODE.NORMAL;
            }


            Logger.ClearProgress();


            // send completed event to MainForm
            UTCheckerEvent args = new UTCheckerEvent();
            args.ReturnCode = this.ReturnCode;
            args.Message = "Completed Event";
            args.Mode = this.Mode;

            OnCompletedEvent(args);

        }

        #endregion



        /// <summary>
        /// This method start the backgroundWorker to work.
        /// </summary>
        public void Run()
        {
            g_bwUTChecker.RunWorkerAsync();
        }


        /// <summary>
        /// Main routine for UTChecker
        /// </summary>
        /// <returns></returns>
        public int RunUTChecker()
        {
            string sFuncName = "[RunUTChecker]";

            string sTDSPath = "";
            string sListFileForTDS = "";

            string sOutputFile = "";

            int dNormalEntryCount = 0;
            int dErrorEntryCount = 0;

            int dTestCaseFuncCount = 0;
            int dRepeatedEntryCount = 0;
            bool bIsErrorEverOccurred = false;
            int dErrorCount = 0;
            int dNGEntryCount = 0;

            int dTestLogIssueCount = 0;
            int dSUTSIssueCount = 0;

            int dByMockito = 0;
            int dByPowerMockito = 0;
            int dVectorcast = 0;

            int dGetterSetter = 0;
            int dEmptymethod = 0;
            int dAbstractmethod = 0;
            int dInterfacemethod = 0;
            int dNativemethod = 0;

            int dBycodeanalysis = 0;
            int dPurefunctioncalls = 0;
            int dPureUIfunctioncalls = 0;

            int dUnknow = 0;

            CultureInfo cultureEN = new CultureInfo("en-US");


            // record the time at Start.
            DateTime l_startTime = DateTime.Now;
            Logger.Print(sFuncName, l_startTime.ToString(cultureEN), Logger.PrintOption.Both);


            // -------------------------------------------------------------------------------
            // update logger / progress
            // -------------------------------------------------------------------------------
            Logger.ClearAll();
            

            if (this.Mode == RunMode.CommandLine)
            {
                Logger.Print("Command Line mode, Turn off Message.", Logger.PrintOption.Logger);
                Logger.Enable = false;
            }


            // -------------------------------------------------------------------------------
            // update logger / progress
            // -------------------------------------------------------------------------------
            Logger.UpdateProgress(2);
            


            // check all of paths.
            if (!CheckEnvironmentSetting())
            {
                Logger.Print(sFuncName, "Path Error", Logger.PrintOption.Both);

                return -1;
               
            }

            // change process log to the path of output
            Logger.PrepareProcessLog(g_sOutputPath);


            // initial all variables
            InitializeVariable();


            // -------------------------------------------------------------------------------
            // update logger / progress
            // -------------------------------------------------------------------------------
            Logger.UpdateProgress(4);



            // Read the module list, where comment/empty lines will be ignored.
            if (!ReadModulesFromListFile(g_sModuleListFile, ref g_lsModules, true, true))
            {
                Logger.Print(sFuncName, "Read DD module list failed.", Logger.PrintOption.Both);

                return -1;
            }


            // -------------------------------------------------------------------------------
            // update logger / progress
            // -------------------------------------------------------------------------------
            Logger.UpdateProgress(6);



            // Read the name of modules from Summary Rereport.
            Dictionary<string, int> a_lsModuleNameInSmmary = ReadAllModuleNamesFromExcel(g_sSummaryReport);
            if (a_lsModuleNameInSmmary == null || a_lsModuleNameInSmmary.Count == 0)
            {
                Logger.Print(sFuncName, "The modules of name can't be found in Summary Template.", Logger.PrintOption.Both);

                return -1;
            }

            // -------------------------------------------------------------------------------
            // update logger / progress
            // -------------------------------------------------------------------------------
            Logger.UpdateProgress(8);


            // prepare summary report
            string sSummaryReportPath = PrepareSummaryReport(g_sSummaryReport, g_sOutputPath);
            Logger.Print($"Total {g_lsModules.Count} module(s) would be checked.", Logger.PrintOption.Both);


            // -------------------------------------------------------------------------------
            // update logger / progress
            // -------------------------------------------------------------------------------
            int diff = 80 / g_lsModules.Count;
            int value = 10;



            try
            {

                foreach (string sItem in g_lsModules)
                {

                    // record start time
                    DateTime _moduleTiemStart = DateTime.Now;


                    // Write the spliter for reading the error log file easily.
                    // (This section must be located behind the remove-error-file section. Otherwise the message will be written to the previous error log file.)

                    Logger.Print("---------------------------------------------------------------", Logger.PrintOption.Both);
                    Logger.Print(sItem, Logger.PrintOption.Both);
                    Logger.Print("---------------------------------------------------------------", Logger.PrintOption.Both);


                    // -------------------------------------------------------------------------------
                    // update logger / progress
                    // -------------------------------------------------------------------------------
                    Logger.UpdateProgress(value);
                    value = value < 90 ? (value + (diff / 2)) : 90;


                    // for _TDS.list
                    sTDSPath = g_sTDSPath + sItem + "\\";
                    sListFileForTDS = sTDSPath + "_TDS.list";

                    // for report
                    sOutputFile = g_sOutputPath + Constants.REPORT_PREFIX + sItem + ".xlsx";


                    // get the file name for Log file and set to Logger
                    g_sErrorLogFile = g_sOutputPath + Constants.REPORT_PREFIX + sItem + ".log";
                    Logger.FileName = g_sErrorLogFile;

                    // Remove old log file.
                    if (File.Exists(g_sErrorLogFile))
                    {
                        File.Delete(g_sErrorLogFile);
                    }


                    #region Reset all of counts

                    // Reset the counters.
                    g_tTestCaseTable.dSourceFileCount = 0;
                    g_tTestCaseTable.dMethodCount = 0;
                    g_tTestCaseTable.dNormalEntryCount = 0;
                    g_tTestCaseTable.dTestCaseFuncCount = 0;
                    g_tTestCaseTable.dErrorCount = 0;
                    g_tTestCaseTable.dRepeatedEntryCount = 0;

                    g_tTestCaseTable.dNormalEntryCount = 0;
                    g_tTestCaseTable.dRepeatedEntryCount = 0;
                    g_tTestCaseTable.dErrorEntryCount = 0;

                    g_tTestCaseTable.dByNACount = 0;
                    g_tTestCaseTable.dByTestScriptCount = 0;
                    g_tTestCaseTable.dByCodeAnalysisCount = 0;
                    g_tTestCaseTable.dByUnknownCount = 0;

                    g_tTestCaseTable.dTestLogIssueCount = 0;
                    g_tTestCaseTable.dSUTSIssueCount = 0;

                    g_tTestCaseTable.ltItems.Clear();

                    // clear all
                    g_tTestCaseTable.stTestTypeStatistic.mockito = 0;
                    g_tTestCaseTable.stTestTypeStatistic.powermockito = 0;
                    g_tTestCaseTable.stTestTypeStatistic.vectorcast = 0;


                    g_tTestCaseTable.stTestTypeStatistic.gettersetter = 0;
                    g_tTestCaseTable.stTestTypeStatistic.emptymethod = 0;
                    g_tTestCaseTable.stTestTypeStatistic.abstractmethod = 0;
                    g_tTestCaseTable.stTestTypeStatistic.interfacemethod = 0;
                    g_tTestCaseTable.stTestTypeStatistic.nativemethod = 0;

                    g_tTestCaseTable.stTestTypeStatistic.codeanalysis = 0;
                    g_tTestCaseTable.stTestTypeStatistic.purefunctioncalls = 0;
                    g_tTestCaseTable.stTestTypeStatistic.pureUIfunctioncalls = 0;


                    g_tTestCaseTable.stTestTypeStatistic.unknow = 0;


                    #endregion



                    try
                    {

                        //
                        // 1. Create a thread to collect Test Logs (By thread)
                        //
                        var testLogItems = new Tuple<string, string>(sItem, g_sTestLogPath);
                        Thread collectThread = new Thread(new ParameterizedThreadStart(CollectTestLogs));
                        collectThread.Priority = ThreadPriority.AboveNormal;
                        //collectTestLogThread.IsBackground = true;
                        collectThread.Start(testLogItems);



                        //
                        // 2. Prepare a dummy output report in case the flow is aborted due to a failure.
                        //
                        if (!PrepareDummayOutputReport(sOutputFile))
                        {
                            bIsErrorEverOccurred = true;
                            Logger.Print($"PrepareDummayOutputReport Fail.", Logger.PrintOption.Both);
                            continue;
                        }


                        //
                        // 3. Collect all TDS files.
                        //
                        g_lsTDSFiles = CollectTDSFiles(sTDSPath, sListFileForTDS);
                        if ((null == g_lsTDSFiles) || (0 == g_lsTDSFiles.Count))
                        {
                            Logger.Print($"TDS file(s) not found.", Logger.PrintOption.Both);
                            continue;
                        }


                        //
                        // 4. Search SUTS
                        //
                        g_sSUTSDocumentPath = SearchSUTSDocumentPath(sItem, g_sSUTSPath);
                        if (g_sSUTSDocumentPath == String.Empty)
                        {
                            Logger.Print($"The SUTS of {sItem} have not found.", Logger.PrintOption.Both);
                        }


                        //
                        // 5. Wait the task "collect test log" done (By thread)
                        //

                        collectThread.Join();


                        //var testLogItems = new Tuple<string, string>(sItem, g_sTestLogPath);
                        //CollectTestLogs(testLogItems);
                        if ((null == g_lsTestLogs) || (0 == g_lsTestLogs.Count))
                        {
                            Logger.Print($"Test log file(s) not found.", Logger.PrintOption.Both);
                        }


                        // -------------------------------------------------------------------------------
                        // update logger / progress
                        // -------------------------------------------------------------------------------
                        Logger.UpdateProgress(value);
                        value = value < 90 ? (value + (diff / 2)) : 90;


                        //
                        // Parse the function & TC info from each TDS file.
                        //
                        if (!ReadDataFromTDSFiles(sItem, ref g_lsTDSFiles, ref g_lsTestLogs, g_sSUTSDocumentPath))
                        {
                            bIsErrorEverOccurred = true;
                            Logger.Print($"Parse TDS file failed, Skipped.", Logger.PrintOption.Both);
                            continue;
                        }



                        // Count designs and methods.
                        CountAndMarkResults();



                        // -------------------------------------------------------------------------------
                        // update logger / progress
                        // -------------------------------------------------------------------------------
                        Logger.Print("Write the result to report.");


                        // Write the results to as an overall lookup table.
                        if (!SaveResults(g_sTemplateFile, sOutputFile, ref g_lsTestLogs))
                        {
                            bIsErrorEverOccurred = true;
                        }



                        int index = -1;
                        if (a_lsModuleNameInSmmary.TryGetValue(sItem.Replace("_", " "), out index))
                        {

                            // Write the overall result of current module
                            if (!WriteSummaryReport(sSummaryReportPath, g_tTestCaseTable, index))
                            {
                                Logger.Print(sFuncName, "Error occurred when writeing data into Summary report.", Logger.PrintOption.Both);
                                bIsErrorEverOccurred = true;
                            }
                        }
                        else
                        {
                            Logger.Print(sFuncName, $"{sItem} has no found in Summary Template", Logger.PrintOption.Both);
                            bIsErrorEverOccurred = true;
                        }



                    }
                    finally
                    {

                        #region Show the result when current module has checked.


                        Logger.Print("", "---------------------------------------------------------------");
                        Logger.Print("", $" {sItem} Module Summary:");
                        Logger.Print("", "---------------------------------------------------------------");


                        // Show the # based on TDS entries. 
                        Logger.Print("  Total # of test cases defined in TDS:", g_tTestCaseTable.ltItems.Count.ToString());
                        Logger.Print("   - Normal entries:  ", String.Format("{0,4}", g_tTestCaseTable.dNormalEntryCount.ToString()));
                        Logger.Print("   - Repeated entries:", String.Format("{0,4}", g_tTestCaseTable.dRepeatedEntryCount.ToString()));
                        Logger.Print("   - Error entries:   ", String.Format("{0,4}", g_tTestCaseTable.dErrorEntryCount.ToString()));
                        int dSum = g_tTestCaseTable.dNormalEntryCount + g_tTestCaseTable.dRepeatedEntryCount + g_tTestCaseTable.dErrorEntryCount;
                        if (dSum != g_tTestCaseTable.ltItems.Count)
                        {
                            Logger.Print("     (Error:", dSum.ToString() + " != " + g_tTestCaseTable.ltItems.Count.ToString() + ")");
                        }

                        // new line
                        Logger.Print("", "");

                        // Show the # based on test means. 
                        Logger.Print("   - Tested by test scripts: ", String.Format("{0,4}", g_tTestCaseTable.dByTestScriptCount.ToString()));
                        Logger.Print("     By Mockito:             ", String.Format("{0,4}", g_tTestCaseTable.stTestTypeStatistic.mockito));
                        Logger.Print("     By PowerMockito:        ", String.Format("{0,4}", g_tTestCaseTable.stTestTypeStatistic.powermockito));
                        Logger.Print("     By VectorCast:          ", String.Format("{0,4}", g_tTestCaseTable.stTestTypeStatistic.vectorcast));
                        Logger.Print("", "");

                        Logger.Print("  Test means summary:", "");
                        Logger.Print("   - No test needed:         ", String.Format("{0,4}", g_tTestCaseTable.dByNACount.ToString()));
                        Logger.Print("     Getter/Setter:          ", String.Format("{0,4}", g_tTestCaseTable.stTestTypeStatistic.gettersetter));
                        Logger.Print("     Empty method:           ", String.Format("{0,4}", g_tTestCaseTable.stTestTypeStatistic.emptymethod));
                        Logger.Print("     Abstract method:        ", String.Format("{0,4}", g_tTestCaseTable.stTestTypeStatistic.abstractmethod));
                        Logger.Print("     Interface method:       ", String.Format("{0,4}", g_tTestCaseTable.stTestTypeStatistic.interfacemethod));
                        Logger.Print("     Native method:          ", String.Format("{0,4}", g_tTestCaseTable.stTestTypeStatistic.nativemethod));
                        Logger.Print("", "");

                        Logger.Print("   - Tested by code analysis:", String.Format("{0,4}", g_tTestCaseTable.dByCodeAnalysisCount.ToString()));
                        Logger.Print("     By code analysis:       ", String.Format("{0,4}", g_tTestCaseTable.stTestTypeStatistic.codeanalysis));
                        Logger.Print("     Pure function calls:    ", String.Format("{0,4}", g_tTestCaseTable.stTestTypeStatistic.purefunctioncalls));
                        Logger.Print("     Pure UI function calls: ", String.Format("{0,4}", g_tTestCaseTable.stTestTypeStatistic.pureUIfunctioncalls));
                        Logger.Print("", "");


                        Logger.Print("   - Unknow items:           ", String.Format("{0,4}", g_tTestCaseTable.dByUnknownCount.ToString()));
                        Logger.Print("     Uknow:                  ", String.Format("{0,4}", g_tTestCaseTable.stTestTypeStatistic.unknow));
                        Logger.Print("", "");


                        dSum = g_tTestCaseTable.dByTestScriptCount + g_tTestCaseTable.dByCodeAnalysisCount +
                            g_tTestCaseTable.dByUnknownCount + g_tTestCaseTable.dByNACount;
                        if (dSum != g_tTestCaseTable.dNormalEntryCount)
                        {
                            Logger.Print("     (Error:", dSum.ToString() + " != " + g_tTestCaseTable.dNormalEntryCount.ToString() + ")");
                        }

                        Logger.Print("  Total # of Test Log Issues:", g_tTestCaseTable.dTestLogIssueCount.ToString());
                        Logger.Print("  Total # of SUTS issues:", g_tTestCaseTable.dSUTSIssueCount.ToString());

                        Logger.Print("  Total # of NG entry(s) found:", g_tTestCaseTable.dNGEntryCount.ToString());


                        #endregion


                        #region Accumulate the counts.

                        // Accumulate the counts.
                        dNormalEntryCount = dNormalEntryCount + g_tTestCaseTable.dNormalEntryCount;
                        dRepeatedEntryCount = dRepeatedEntryCount + g_tTestCaseTable.dRepeatedEntryCount;
                        dErrorEntryCount = dErrorEntryCount + g_tTestCaseTable.dErrorEntryCount;

                        dTestCaseFuncCount = dTestCaseFuncCount + g_tTestCaseTable.dTestCaseFuncCount;
                        dErrorCount = dErrorCount + g_tTestCaseTable.dErrorCount;
                        dNGEntryCount = dNGEntryCount + g_tTestCaseTable.dNGEntryCount;

                        dTestLogIssueCount = dTestLogIssueCount + g_tTestCaseTable.dTestLogIssueCount;
                        dSUTSIssueCount = dSUTSIssueCount + g_tTestCaseTable.dSUTSIssueCount;

                        dByMockito = dByMockito + g_tTestCaseTable.stTestTypeStatistic.mockito;
                        dByPowerMockito = dByPowerMockito + g_tTestCaseTable.stTestTypeStatistic.powermockito;
                        dVectorcast = dVectorcast + g_tTestCaseTable.stTestTypeStatistic.vectorcast;

                        dGetterSetter = dGetterSetter + g_tTestCaseTable.stTestTypeStatistic.gettersetter;
                        dEmptymethod = dEmptymethod + g_tTestCaseTable.stTestTypeStatistic.emptymethod;
                        dAbstractmethod = dAbstractmethod + g_tTestCaseTable.stTestTypeStatistic.abstractmethod;
                        dInterfacemethod = dInterfacemethod + g_tTestCaseTable.stTestTypeStatistic.interfacemethod;
                        dNativemethod = dNativemethod + g_tTestCaseTable.stTestTypeStatistic.nativemethod;

                        dBycodeanalysis = dBycodeanalysis + g_tTestCaseTable.stTestTypeStatistic.codeanalysis;
                        dPurefunctioncalls = dPurefunctioncalls + g_tTestCaseTable.stTestTypeStatistic.purefunctioncalls;
                        dPureUIfunctioncalls = dPureUIfunctioncalls + g_tTestCaseTable.stTestTypeStatistic.pureUIfunctioncalls;
                        dUnknow = dUnknow + g_tTestCaseTable.stTestTypeStatistic.unknow;

                        #endregion


                    }


                    // record elapsed time for each module.
                    DateTime _moduleTiemEnd = DateTime.Now;
                    TimeSpan intv = _moduleTiemEnd - _moduleTiemStart;
                    Logger.Print($"Elapsed Time: {intv.ToString()}", Logger.PrintOption.Both);



                } // End of foreach

                // -------------------------------------------------------------------------------
                // update logger / progress
                // -------------------------------------------------------------------------------
                Logger.UpdateProgress(90);

            }
            finally
            {

                #region Show Overall Summary Info

                // Show overall summary info.
                Logger.Print("", "---------------------------------------------------------------", Logger.PrintOption.Both);
                Logger.Print("", " Overall Summary:", Logger.PrintOption.Both);
                Logger.Print("", "---------------------------------------------------------------", Logger.PrintOption.Both);

                Logger.Print(" Total # of test cases with repeated labels: " + dRepeatedEntryCount.ToString(), "", Logger.PrintOption.Both);
                Logger.Print(" Total # of non-repeated test cases collected: " + dNormalEntryCount.ToString(), "", Logger.PrintOption.Both);
                Logger.Print(" Total # of non-repeated test case functions collected: " + dTestCaseFuncCount.ToString(), "", Logger.PrintOption.Both);
                Logger.Print(" Total # of errors found: " + dErrorCount.ToString(), "", Logger.PrintOption.Both);

                Logger.Print(" Total # of Test Log Issues found:  " + dTestLogIssueCount.ToString(), "", Logger.PrintOption.Both);
                Logger.Print(" Total # of SUTS Issues found:      " + dSUTSIssueCount.ToString(), "", Logger.PrintOption.Both);
                Logger.Print(" Total # of NG entries found:       " + dNGEntryCount.ToString(), "", Logger.PrintOption.Both);


                Logger.Print(" Total # of By Mockito:             ", String.Format("{0,4}", dByMockito), Logger.PrintOption.Both);
                Logger.Print(" Total # of By PowerMockito:        ", String.Format("{0,4}", dByPowerMockito), Logger.PrintOption.Both);
                Logger.Print(" Total # of By VectorCast:          ", String.Format("{0,4}", dVectorcast), Logger.PrintOption.Both);


                Logger.Print(" Total # of Getter/Setter:          ", String.Format("{0,4}", dGetterSetter), Logger.PrintOption.Both);
                Logger.Print(" Total # of Empty method:           ", String.Format("{0,4}", dEmptymethod), Logger.PrintOption.Both);
                Logger.Print(" Total # of Abstract method:        ", String.Format("{0,4}", dAbstractmethod), Logger.PrintOption.Both);
                Logger.Print(" Total # of Interface method:       ", String.Format("{0,4}", dInterfacemethod), Logger.PrintOption.Both);
                Logger.Print(" Total # of Native method:          ", String.Format("{0,4}", dNativemethod), Logger.PrintOption.Both);
                Logger.Print(" Total # of By code analysis:       ", String.Format("{0,4}", dBycodeanalysis), Logger.PrintOption.Both);
                Logger.Print(" Total # of Pure function calls:    ", String.Format("{0,4}", dPurefunctioncalls), Logger.PrintOption.Both);
                Logger.Print(" Total # of Pure UI function calls: ", String.Format("{0,4}", dPureUIfunctioncalls), Logger.PrintOption.Both);
                Logger.Print(" Total # of Uknow:                  ", String.Format("{0,4}", dUnknow), Logger.PrintOption.Both);
                Logger.Print("", "");


                #endregion





                // release Office.
                ReleaseOfficeApps();



                // Show ending message.
                if (bIsErrorEverOccurred)
                {
                    Logger.Print(sFuncName, "Failed!", Logger.PrintOption.Both);
                }
                else
                {

                    //Logger.Print("Update path setting to " + UTCheckerSetting.FileName);
                    //WriteEnvironmentSetting();

                }
            }

            // record the time at Finish
            DateTime l_finishTime = DateTime.Now;

            // update looger
            Logger.Print("All Jobs Done! " + l_finishTime.ToString(cultureEN), Logger.PrintOption.Both);

            // print Elapsed Time
            TimeSpan interval = l_finishTime - l_startTime;
            Logger.Print($"Elapsed Time: {interval.Hours}:{interval.Minutes}:{interval.Seconds}:{interval.Milliseconds}", Logger.PrintOption.Both);



            // -------------------------------------------------------------------------------
            // update logger / progress
            // -------------------------------------------------------------------------------
            Logger.UpdateProgress(100);

            return 0;
        }




        /// <summary>
        /// Init all variables
        /// </summary>
        /// <returns></returns>
        private bool InitializeVariable()
        {
            string sFuncName = "[InitializeVariable]";


            try
            {
                foreach (Process proc in Process.GetProcessesByName("EXCEL"))
                {
                    proc.Kill();
                }

                foreach (Process proc in Process.GetProcessesByName("WINWORD"))
                {
                    proc.Kill();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }



            // get the handler for excel application
            g_excelApp = new Excel.Application
            {
                DisplayAlerts = false
            };

            // get the handler for word application
            g_wordApp = new Word.Application();
            g_wordApp.DisplayAlerts = Microsoft.Office.Interop.Word.WdAlertLevel.wdAlertsNone;

            // init the list of modules
            g_lsModules = new List<string>();

            // Initialize a list for TDS files
            g_lsTDSFiles = new List<string>();

            // New the list for storing the data read.
            g_tTestCaseTable.ltItems = new List<TestCaseItem>();


            Logger.Print(sFuncName, "Done", Logger.PrintOption.Both);


            return true;
        }



        public bool UpdateEnvironmentSetting(EnvrionmentSetting env)
        {

            g_FilePathSetting = env;

            g_sModuleListFile = env.listFile;           // list file
            g_sTDSPath = env.tdsPath;                    // tds path
            g_sOutputPath = env.outputPath;              // output path
            g_sTemplateFile = env.reportTemplate; ; ;    // template file
            g_sSummaryReport = env.summaryTemplate;      // summary templat
            g_sTestLogPath = env.testlogPath;
            g_sSUTSPath = env.sutsPath;
            g_sReferenceListsPath = env.referenceListsPath; // reference lists path
            g_sSUTRRPath = env.sutrrPath;

            this.Mode = RunMode.User;

            return true;

        }


        /// <summary>
        /// Update the setting of environment from Command to MainFrom, and Mainform to Globals
        /// </summary>
        /// <param name="commandline"></param>
        /// <returns></returns>
        public bool UpdateEnvironmentSetting(string[] args)
        {
            string sFuncName = "[UpdateEnvironmentSetting]";

            bool bDone = false;

            EnvrionmentSetting ps = new EnvrionmentSetting();

            //char[] delimiter = new char[] { ' ', '"' };
            //string[] args = Environment.CommandLine.Split(delimiter, StringSplitOptions.RemoveEmptyEntries);

            if (args.Length == Constants.CommandArguments.Match)
            {

                // update mode
                this.Mode = RunMode.CommandLine;

                //
                // update setting to textbox
                //
                ps.listFile = args[1];
                ps.tdsPath = args[2];
                ps.outputPath = args[3];
                ps.reportTemplate = args[4];
                ps.summaryTemplate = args[5];
                ps.testlogPath = args[6];
                ps.sutsPath = args[7];

                // update to global variable
                this.g_FilePathSetting = ps;

                // event
                UTCheckerEvent eventArgs = new UTCheckerEvent();
                eventArgs.Path = this.g_FilePathSetting;
                eventArgs.Mode = this.Mode;
                OnUpdatePathEvent(eventArgs);

                bDone = true;

                Logger.Print(sFuncName, "Update the path setting from Command Line.");

            }
            else
            {


                // update mode
                this.Mode = RunMode.User;

                if (File.Exists(UTCheckerSetting.FileName))
                {
                    string[] lines = System.IO.File.ReadAllLines(@"UTChecker.setting");

                    foreach (string s in lines)
                    {
                        if (s != "")
                        {
                            char[] sper = new char[] { ' ', '=' };
                            string[] setting = s.Split(sper, StringSplitOptions.RemoveEmptyEntries);

                            if (setting[0].Equals(UTCheckerSetting.Prefix))
                            {

                                if (setting[1].Equals(UTCheckerSetting.ListFile))
                                {
                                    ps.listFile = setting[2];
                                }
                                else if (setting[1].Equals(UTCheckerSetting.TDSPath))
                                {
                                    ps.tdsPath = setting[2];
                                }
                                else if (setting[1].Equals(UTCheckerSetting.OutputPath))
                                {
                                    ps.outputPath = setting[2];
                                }
                                else if (setting[1].Equals(UTCheckerSetting.ReportTemplate))
                                {
                                    ps.reportTemplate = setting[2];
                                }
                                else if (setting[1].Equals(UTCheckerSetting.SummaryTemplate))
                                {
                                    ps.summaryTemplate = setting[2];
                                }
                                else if (setting[1].Equals(UTCheckerSetting.TestLogPath))
                                {
                                    ps.testlogPath = setting[2];
                                }
                                else if (setting[1].Equals(UTCheckerSetting.ReferenceListsPath))
                                {
                                    ps.referenceListsPath = setting[2];
                                }
                                else if (setting[1].Equals(UTCheckerSetting.SUTS_PATH))
                                {
                                    ps.sutsPath = setting[2];
                                }
                                else if (setting[1].Equals(UTCheckerSetting.SURR_PATH))
                                {
                                    ps.sutrrPath = setting[2];
                                }
                                else
                                {

                                }

                            }
                        }
                    }

                    g_FilePathSetting = ps;

                    // event
                    UTCheckerEvent eventArgs = new UTCheckerEvent();
                    eventArgs.Path = this.g_FilePathSetting;
                    eventArgs.Mode = this.Mode;
                    OnUpdatePathEvent(eventArgs);


                    bDone = true;

                    Logger.Print(sFuncName, "Update the path setting from UTChecker.setting");

                }
                else
                {
                    bDone = true;

                    Logger.Print(sFuncName, "The path settings haven't been updated.");

                }

            }


            return bDone;
        }



        /// <summary>
        /// Check the setting of enviroment before starting to check the UT 
        /// </summary>
        /// <returns></returns>
        private bool CheckEnvironmentSetting()
        {
            string sFuncName = "[CheckEnvironmentSetting]";


            EnvrionmentSetting fp = g_FilePathSetting;

            g_sModuleListFile =  fp.listFile;           // list file
            g_sTDSPath = fp.tdsPath;                    // tds path
            g_sOutputPath = fp.outputPath;              // output path
            g_sTemplateFile = fp.reportTemplate; ; ;    // template file
            g_sSummaryReport = fp.summaryTemplate;      // summary templat
            g_sTestLogPath = fp.testlogPath;
            g_sSUTSPath = fp.sutsPath;
            g_sReferenceListsPath = fp.referenceListsPath; // reference lists path
            g_sSUTRRPath = fp.sutrrPath;


            // Ensure each path is ended with a '\\'.
            if (!g_sTDSPath.EndsWith("\\"))
            {
                g_sTDSPath = g_sTDSPath + "\\";
            }

            if (!g_sOutputPath.EndsWith("\\"))
            {
                g_sOutputPath = g_sOutputPath + "\\";
            }

            if (!g_sTestLogPath.EndsWith("\\"))
            {
                g_sTestLogPath = g_sTestLogPath + "\\";
            }


            // Check the existence of the input & output paths.
            if (!Directory.Exists(g_sTDSPath))
            {
                Logger.Print(sFuncName, "Cannot find the path of TDS: " + g_sTDSPath);
                return false;
            }

            // check the path of test logs.
            if (!Directory.Exists(g_sTestLogPath))
            {
                Logger.Print(sFuncName, "Cannot find the path of Test Log: " + g_sTestLogPath);
                return false;
            }


            // check the path of SUTS
            if (!Directory.Exists(g_sSUTSPath))
            {
                Logger.Print(sFuncName, "Cannot find the path of SUTS: " + g_sSUTSPath);
                return false;
            }

            // check the path of output
            if (!Directory.Exists(g_sOutputPath))
            {
                Logger.Print(sFuncName, "Cannot find the path for output: " + g_sTestLogPath);
                return false;
            }



            //// check the path of reference
            //if (!Directory.Exists(g_sReferenceListsPath))
            //{
            //    Logger.Print(sFuncName, "Cannot find the reference lists path: " + g_sReferenceListsPath);
            //    return false;
            //}

            //// check the path of SUTRR
            //if (!Directory.Exists(g_sSUTRRPath))
            //{
            //    Logger.Print(sFuncName, "Cannot find the SUTRR path: " + g_sSUTRRPath);
            //    return false;
            //}



            // Update the input/output files, if needs.
            if (!g_sModuleListFile.Contains("\\"))
            {
                g_sModuleListFile = g_sOutputPath + g_sModuleListFile;
            }
            if (!g_sTemplateFile.Contains("\\"))
            {
                g_sTemplateFile = g_sOutputPath + g_sTemplateFile;
            }

            // Check the existence of the input files.
            if (!File.Exists(g_sModuleListFile))
            {
                Logger.Print(sFuncName, "Cannot find Module List File: " + g_sModuleListFile);
                return false;
            }
            if (!File.Exists(g_sTemplateFile))
            {
                Logger.Print(sFuncName, "Cannot find the template file: " + g_sTemplateFile);
                return false;
            }

            if (!File.Exists(g_sSummaryReport))
            {
                Logger.Print(sFuncName, "Cannot find the summary report: " + g_sSummaryReport);
               return false;
            }


            Logger.Print(sFuncName, "Done", Logger.PrintOption.Both);

            return true;
        }



        /// <summary>
        /// Write the setting of environment to file.
        /// </summary>
        /// <returns></returns>
        private bool WriteEnvironmentSetting()
        {
            //string sFuncName = "[WriteEnvironmentSetting]";

            if (this.Mode == RunMode.CommandLine)
            {
                return false;
            }
            else
            {

                if (File.Exists(UTCheckerSetting.FileName))
                {
                    File.Delete(UTCheckerSetting.FileName);
                }

                using (System.IO.StreamWriter file = new System.IO.StreamWriter(UTCheckerSetting.FileName))
                {

                    file.WriteLine(UTCheckerSetting.Prefix + " " +
                                    UTCheckerSetting.ListFile + "=" +
                                    g_FilePathSetting.listFile);

                    file.WriteLine(UTCheckerSetting.Prefix + " " +
                                    UTCheckerSetting.TDSPath + "=" +
                                    g_FilePathSetting.tdsPath);

                    file.WriteLine(UTCheckerSetting.Prefix + " " +
                                    UTCheckerSetting.OutputPath + "=" +
                                    g_FilePathSetting.outputPath);

                    file.WriteLine(UTCheckerSetting.Prefix + " " +
                                    UTCheckerSetting.ReportTemplate + "=" +
                                    g_FilePathSetting.reportTemplate);


                    file.WriteLine(UTCheckerSetting.Prefix + " " +
                                    UTCheckerSetting.SummaryTemplate + "=" +
                                    g_FilePathSetting.summaryTemplate);


                    file.WriteLine(UTCheckerSetting.Prefix + " " +
                                    UTCheckerSetting.TestLogPath + "=" +
                                    g_FilePathSetting.testlogPath);

                    file.WriteLine(UTCheckerSetting.Prefix + " " +
                                    UTCheckerSetting.SUTS_PATH + "=" +
                                    g_FilePathSetting.sutsPath);

                }

            }


            return true;
        }






        /// <summary>
        /// Read the list of Modules from a List File.
        /// </summary>
        /// <param name="a_sInFile"></param>
        /// <param name="a_lsOutList"></param>
        /// <param name="a_bTrim"></param>
        /// <param name="a_bValidLinesOnly"></param>
        /// <returns></returns>
        private bool ReadModulesFromListFile(string a_sInFile, ref List<string> a_lsOutList, bool a_bTrim, bool a_bValidLinesOnly)
        {
            string sFuncName = "[ReadTextFileToStringList]";

            // Check the existence of the input file.
            if (!File.Exists(a_sInFile))
            {
                Logger.Print(sFuncName, "Cannot find \"" + a_sInFile + "\"");
                return false;
            }

            // New a list if...
            if (null == a_lsOutList)
            {
                a_lsOutList = new List<string>();
            }


            try
            {
                // Clear the buffer first.
                a_lsOutList.Clear();

                // Read lines from the text file.
                string[] sLines = File.ReadAllLines(a_sInFile);

                // Assign the read lines to the list box.
                if (0 < sLines.Length)
                {
                    // Assign the read lines to the list box.
                    a_lsOutList.AddRange(sLines);

                    if (a_bTrim)
                    {
                        for (int i = 0; i < a_lsOutList.Count; i++)
                        {
                            a_lsOutList[i] = a_lsOutList[i].Trim();
                        }
                    }

                    // Remove invalid lines from the list:
                    // (1) Comment lines (line starts with a '#')
                    // (2) Empty lines
                    // (3) Lines contain spaces only
                    if (a_bValidLinesOnly)
                    {
                        for (int i = a_lsOutList.Count - 1; i >= 0; i--)
                        {
                            string sLine = a_lsOutList[i];

                            // Remove the line start with "#".
                            if (sLine.StartsWith("#"))
                            {
                                a_lsOutList.RemoveAt(i);
                                continue;
                            }

                            // Remove empty or space line.
                            if ("" == sLine.Replace(" ", ""))
                            {
                                a_lsOutList.RemoveAt(i);
                            }
                        }
                    }
                }

                if (0 == a_lsOutList.Count)
                {
                    Logger.Print(sFuncName, "No-line is loaded from " + a_sInFile);
                }


            }
            catch (Exception ex)
            {
                Logger.Print(sFuncName, ex.ToString());
                return false;
            }

            Logger.Print(sFuncName, "Done", Logger.PrintOption.Both);

            return true;
        }







        /// <summary>
        /// 
        /// </summary>
        /// <param name="a_sOutputFile"></param>
        /// <returns></returns>
        private bool PrepareDummayOutputReport(string a_sOutputFile)
        {
            string sFuncName = "[PrepareDummayOutputReport]";

            try
            {
                // Remove old output file.
                if (File.Exists(a_sOutputFile))
                {
                    File.Delete(a_sOutputFile);
                }

                // Make a copy of the template file as the dummy output file.
                if (!File.Exists(g_sTemplateFile))
                {
                    Logger.Print(sFuncName, "Cannot find " + g_sTemplateFile);
                    return false;
                }

                File.Copy(g_sTemplateFile, a_sOutputFile);

            }
            catch (Exception e)
            {
                Logger.Print(sFuncName, "Exception: " + e.ToString());
                return false;
            }

            Logger.Print(sFuncName, "Done", Logger.PrintOption.Both);

            return true;
        }






        class SourceFileNameCompare : IEqualityComparer<TestCaseItem>
        {
            #region IEqualityComparer<Person> Members

            public bool Equals(TestCaseItem x, TestCaseItem y)
            {
                return x.sSourceFileName.Equals(y.sSourceFileName);
            }

            public int GetHashCode(TestCaseItem obj)
            {
                return obj.sSourceFileName.GetHashCode();
            }

            #endregion
        }

        class MethodNameCompare : IEqualityComparer<TestCaseItem>
        {
            #region IEqualityComparer<Person> Members

            public bool Equals(TestCaseItem x, TestCaseItem y)
            {
                return x.sMethodName.Equals(y.sMethodName);
            }

            public int GetHashCode(TestCaseItem obj)
            {
                return obj.sMethodName.GetHashCode();
            }

            #endregion
        }

        class TestCaseFuncCompare : IEqualityComparer<TestCaseItem>
        {
            #region IEqualityComparer<Person> Members

            public bool Equals(TestCaseItem x, TestCaseItem y)
            {
                return x.sTCFuncName.Equals(y.sTCFuncName);
            }

            public int GetHashCode(TestCaseItem obj)
            {
                return obj.sTCFuncName.GetHashCode();
            }

            #endregion
        }


        /// <summary>
        /// 
        /// </summary>
        public void CountAndMarkResults()
        {
            const string sFuncName = "[CountAndMarkResults]";

            int dErrorEntryCount = 0;
            int dRepeatedEntryCount = 0;
            int dNormalEntryCount = 0;

            int dDoneByNACount = 0;
            int dDoneByCodeAnalysisCount = 0;
            int dDoneByTestScriptCount = 0;
            int dDoneByOthersCount = 0;

            // fo statistic
            int dByMockito = 0;
            int dByPowerMockito = 0;
            int dGetterSetter = 0;
            int dEmptymethod = 0;
            int dAbstractmethod = 0;
            int dInterfacemethod = 0;
            int dNativemethod = 0;
            int dBycodeanalysis = 0;
            int dPurefunctioncalls = 0;
            int dPureUIfunctioncalls = 0;
            int dVectorCast = 0;
            int dUnknow = 0;


            try
            {
                // -------------------------------------------------------
                // Count the # of non-repeated source files.
                // -------------------------------------------------------
                // Form a list of non-repeated test cases.
                List<TestCaseItem> ltNonRepeatedItems = g_tTestCaseTable.ltItems.Distinct(new SourceFileNameCompare()).ToList(); ;

                // Remove N/A & error entries.
                for (int i = ltNonRepeatedItems.Count - 1; i >= 0; i--)
                {
                    if (ltNonRepeatedItems[i].sSourceFileName.StartsWith(Constants.StringTokens.NA) ||
                        ltNonRepeatedItems[i].sSourceFileName.StartsWith(Constants.StringTokens.ERROR_MSG_HEADER))
                    {
                        ltNonRepeatedItems.RemoveAt(i);
                    }
                }

                // Record the number of non-repeated source files.
                g_tTestCaseTable.dSourceFileCount = ltNonRepeatedItems.Count;

                // -------------------------------------------------------
                // Count the # of non-repeated methods.
                // -------------------------------------------------------
                // Form a list of non-repeated test cases.
                ltNonRepeatedItems = g_tTestCaseTable.ltItems.Distinct(new MethodNameCompare()).ToList(); ;

                // Remove N/A & error entries.
                for (int i = ltNonRepeatedItems.Count - 1; i >= 0; i--)
                {
                    if (ltNonRepeatedItems[i].sMethodName.StartsWith(Constants.StringTokens.NA) ||
                        ltNonRepeatedItems[i].sMethodName.StartsWith(Constants.StringTokens.ERROR_MSG_HEADER))
                    {
                        ltNonRepeatedItems.RemoveAt(i);
                    }
                }

                // Record the number of non-repeated source files.
                g_tTestCaseTable.dMethodCount = ltNonRepeatedItems.Count;

                // -------------------------------------------------------
                // Count the # of non-repeated TC functions to be implemented.
                // -------------------------------------------------------
                // Form a list of non-repeated test case functions.
                ltNonRepeatedItems = g_tTestCaseTable.ltItems.Distinct(new TestCaseFuncCompare()).ToList(); ;

                // Remove N/A & error entries.
                for (int i = ltNonRepeatedItems.Count - 1; i >= 0; i--)
                {
                    if (ltNonRepeatedItems[i].sTCFuncName.StartsWith(Constants.StringTokens.NA) ||
                        ltNonRepeatedItems[i].sTCFuncName.StartsWith(Constants.StringTokens.ERROR_MSG_HEADER))
                    {
                        ltNonRepeatedItems.RemoveAt(i);
                    }
                }

                // Record the counted result.
                g_tTestCaseTable.dTestCaseFuncCount = ltNonRepeatedItems.Count;

                // -------------------------------------------------------
                // Count the # of TC labels, error entries, and testing means.
                // -------------------------------------------------------
                // Sort test cases.
                g_tTestCaseTable.ltItems = g_tTestCaseTable.ltItems.OrderBy(x => x.sTCLabelName).ToList();

                // Check N/A entries & duplicate entries which have the same test case label.
                TestCaseItem tTestCase;
                dErrorEntryCount = 0;
                dRepeatedEntryCount = 0;
                dNormalEntryCount = 0;

                for (int i = g_tTestCaseTable.ltItems.Count - 1; i >= 0; i--)
                {
                    tTestCase = g_tTestCaseTable.ltItems[i];

                    // Case 1: Error TC label (including "N/A" ones, as "N/A" is not allowed)
                    if (tTestCase.sTCLabelName.StartsWith(Constants.StringTokens.ERROR_MSG_HEADER))
                    {
                        dErrorEntryCount++;
                        continue;
                    }

                    // Case 2: Repeated TC label
                    else if ((0 < i) && (tTestCase.sTCLabelName == g_tTestCaseTable.ltItems[i - 1].sTCLabelName))
                    {
                        dRepeatedEntryCount++;

                        // Check if the repeated TCs are testing the same method.
                        // Case 2a: Yes --> Duplicate entry found
                        if (tTestCase.sMethodName == g_tTestCaseTable.ltItems[i - 1].sMethodName)
                        {
                            // Update the flag.
                            tTestCase.bIsRepeated = true;
                            tTestCase.sTCLabelName = Constants.StringTokens.ERROR_MSG_HEADER + ErrorMessage.DUPLICATE_TC_LABEL_FOUND + ": \"" + tTestCase.sTCLabelName + "\"";

                            // Flush back the item.
                            g_tTestCaseTable.ltItems[i] = tTestCase;
                        }
                        // Case 2b: No --> 1 TC for multiple methods
                        // Allowed case. Do nothing.
                    }

                    // Case 3: Normal
                    else
                    {
                        dNormalEntryCount++;

                        // Count the test means.
                        if (TestMeans.NA == tTestCase.eTestMeans)
                        {
                            dDoneByNACount++;
                        }
                        else if (TestMeans.TEST_SCRIPT == tTestCase.eTestMeans)
                        {
                            dDoneByTestScriptCount++;
                        }
                        else if (TestMeans.CODE_ANALYSIS == tTestCase.eTestMeans)
                        {
                            dDoneByCodeAnalysisCount++;
                        }
                        else
                        {
                            dDoneByOthersCount++;
                        }



                        switch (tTestCase.eType)
                        {
                            case TestType.ByMockito:        dByMockito++;       break;
                            case TestType.ByPowerMockito:   dByPowerMockito++;  break;
                            case TestType.ByVectorCast:     dVectorCast++;      break;


                            case TestType.GetterSetter:     dGetterSetter++;    break;
                            case TestType.Empty:            dEmptymethod++;     break;
                            case TestType.Abstract:         dAbstractmethod++;  break;
                            case TestType.Interface:        dInterfacemethod++; break;
                            case TestType.Native:           dNativemethod++;    break;


                            case TestType.ByCodeAnalysis:
                                dBycodeanalysis++;
                                break;

                            case TestType.PureFunctionCalls:
                                dPurefunctioncalls++;
                                break;

                            case TestType.PureUIFunctionCalls:
                                dPureUIfunctioncalls++;
                                break;


                            default:
                                dUnknow++;
                                break;
                        }

                    }
                }

                // record the statistic of each test type.
                g_tTestCaseTable.dNormalEntryCount = dNormalEntryCount;
                g_tTestCaseTable.dRepeatedEntryCount = dRepeatedEntryCount;
                g_tTestCaseTable.dErrorEntryCount = dErrorEntryCount;

                g_tTestCaseTable.dByNACount = dDoneByNACount;
                g_tTestCaseTable.dByTestScriptCount = dDoneByTestScriptCount;
                g_tTestCaseTable.dByCodeAnalysisCount = dDoneByCodeAnalysisCount;
                g_tTestCaseTable.dByUnknownCount = dDoneByOthersCount;

                g_tTestCaseTable.stTestTypeStatistic.mockito = dByMockito;
                g_tTestCaseTable.stTestTypeStatistic.powermockito = dByPowerMockito;
                g_tTestCaseTable.stTestTypeStatistic.vectorcast = dVectorCast;

                g_tTestCaseTable.stTestTypeStatistic.gettersetter = dGetterSetter;
                g_tTestCaseTable.stTestTypeStatistic.emptymethod = dEmptymethod;
                g_tTestCaseTable.stTestTypeStatistic.abstractmethod = dAbstractmethod;
                g_tTestCaseTable.stTestTypeStatistic.interfacemethod = dInterfacemethod;
                g_tTestCaseTable.stTestTypeStatistic.nativemethod = dNativemethod;

                g_tTestCaseTable.stTestTypeStatistic.codeanalysis = dBycodeanalysis;
                g_tTestCaseTable.stTestTypeStatistic.purefunctioncalls = dPurefunctioncalls;
                g_tTestCaseTable.stTestTypeStatistic.pureUIfunctioncalls = dPureUIfunctioncalls;


                g_tTestCaseTable.stTestTypeStatistic.unknow = dUnknow;


                // Double check the sum of the counts.
                int dSum = g_tTestCaseTable.dNormalEntryCount + g_tTestCaseTable.dRepeatedEntryCount + g_tTestCaseTable.dErrorEntryCount;
                if (dSum != g_tTestCaseTable.ltItems.Count)
                {
                    Logger.Print("     (Error:", dSum.ToString() + " != " + g_tTestCaseTable.ltItems.Count.ToString() + ")");
                }

                dSum = g_tTestCaseTable.dByNACount + g_tTestCaseTable.dByTestScriptCount +
                        g_tTestCaseTable.dByCodeAnalysisCount + g_tTestCaseTable.dByUnknownCount;
                if (dSum != g_tTestCaseTable.dNormalEntryCount)
                {
                    Logger.Print("     (Error:", dSum.ToString() + " != " + g_tTestCaseTable.dNormalEntryCount.ToString() + ")");
                }
            }
            catch (SystemException e)
            {
                Logger.Print(sFuncName, e.ToString());
            }
        }
    }
}
