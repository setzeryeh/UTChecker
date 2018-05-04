using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace UTChecker
{
    public partial class UTChecker
    {
        private MainForm g_MF = null;

        public TimeSpan TimeSpan { get; private set; }

        /// <summary>
        /// Constructor for UTChecker
        /// </summary>
        public UTChecker(MainForm mf)
        {

            g_MF = mf;

            g_FilePathSetting = new EnvrionmentSetting();

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

            g_bwUTChecker.WorkerReportsProgress = true;
            g_bwUTChecker.WorkerSupportsCancellation = true;
            g_bwUTChecker.DoWork += new DoWorkEventHandler(bwUTChecker_DoWork);
            g_bwUTChecker.ProgressChanged += new ProgressChangedEventHandler(bwUTChecker_ProgressChanged);
            g_bwUTChecker.RunWorkerCompleted += new RunWorkerCompletedEventHandler(bwUTChecker_RunWorkerCompleted);
        }

        /// <summary>
        /// This method start the backgroundWorker to work.
        /// </summary>
        public void Run()
        {
            g_bwUTChecker.RunWorkerAsync();
        }


        /// <summary>
        /// An event which triggers the RunTDSParser
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void bwUTChecker_DoWork(object sender, DoWorkEventArgs e)
        {

            // Get the BackgroundWorker that raised this event.
            BackgroundWorker worker = sender as BackgroundWorker;
            e.Result = RunUTChecker();
        }


        /// <summary>
        /// bwUTChecker_ProgressChanged
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void bwUTChecker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            //this.progressBarUTChecker.Value = e.ProgressPercentage;
        }


        /// <summary>
        /// bwUTChecker_RunWorkerCompleted
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void bwUTChecker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (this.RunUTCheckerBy == RunBy.CommandLine)
            {
                Environment.ExitCode = 0;
                Environment.Exit(Environment.ExitCode);
            }
            else
            {
                MessageBox.Show("Done");
                Logger.ClearProgress();
            }
        }



        #endregion


        
        /// <summary>
        /// Main routine for UTChecker
        /// </summary>
        /// <returns></returns>
        public int RunUTChecker()
        {
            string sFuncName = "[TDS_Parser]";

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


            int dByMockito = 0;
            int dByPowerMockito = 0;
            int dBycodeanalysis = 0;
            int dGetterSetter = 0;
            int dEmptymethod = 0;
            int dAbstractmethod = 0;
            int dInterfacemethod = 0;
            int dNativemethod = 0;
            int dPurefunctioncalls = 0;
            int dPureUIfunctioncalls = 0;
            int dUnknow = 0;


            DateTime startTime = DateTime.Now;
            Logger.Print("Run UT checker" + startTime.ToString(new CultureInfo("en-US")), Logger.PrintOption.Both);

            // update logger
            Logger.ClearAll();

            // initial all variables
            InitializeVariable();
            
           
            if (!CheckSetting())
            {
                if (this.RunUTCheckerBy == UTChecker.RunBy.CommandLine)
                {
                    Environment.ExitCode = -1;
                    Environment.Exit(Environment.ExitCode);
                }
                else
                {
                    
                }
                 
            }

            // update logger
            Logger.Print("Initialize variables done.");


            // Read the module list, where comment/empty lines will be ignored.
            if (!ReadModulesFromListFile(g_sModuleListFile, ref g_lsModules, true, true))
            {
                Logger.Print(sFuncName, "Read DD module list failed.", Logger.PrintOption.Both);

                return 0;
            }


            // Read the name of modules from Summary Rereport.
            List<string> nameInSummary = ReadAllModuleNamesFromExcel(g_sSummaryReport);


            // update logger
            Logger.Print("Read List file done.");


            // prepare summary report
            string sSummaryReportPath = PrepareSummaryReport(g_sSummaryReport, g_sOutputPath);


            // update 
            int diff = 80 / g_lsModules.Count;
            int value = 10;
            Logger.UpdateProgress(value);


            Logger.Print($"Total {g_lsModules.Count} module(s) would be checked.", Logger.PrintOption.Both);


            try
            {

                foreach (string sItem in g_lsModules)
                {

                    // for _TDS.list
                    sTDSPath = g_sTDSPath + sItem + "\\";
                    sListFileForTDS = sTDSPath + "_TDS.list";

                    // for report
                    sOutputFile = g_sOutputPath + Constants.REPORT_PREFIX + sItem + ".xlsx";



                    // for Error log
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

                    g_tTestCaseTable.ltItems.Clear();

                    // clear all
                    gn_ByMockito = 0;
                    gn_ByPowerMockito = 0;
                    gn_Bycodeanalysis = 0;
                    gn_GetterSetter = 0;
                    gn_Emptymethod = 0;
                    gn_Abstractmethod = 0;
                    gn_Interfacemethod = 0;
                    gn_Nativemethod = 0;
                    gn_Purefunctioncalls = 0;
                    gn_PureUIfunctioncalls = 0;
                    gn_Unknow = 0;

                    #endregion


                    // Write the spliter for reading the error log file easily.
                    // (This section must be located behind the remove-error-file section. Otherwise the message will be written to the previous error log file.)
                    Logger.Print("", "---------------------------------------------------------------", Logger.PrintOption.Both);
                    Logger.Print("", sItem, Logger.PrintOption.Both);
                    Logger.Print("", "---------------------------------------------------------------", Logger.PrintOption.Both);

                   
                    try
                    {
                        // update logger
                        Logger.Print("Prepare a output report.");

                        // Prepare a dummy output report in case the flow is aborted due to a failure.
                        if (!PrepareDummayOutputReport(sOutputFile))
                        {
                            bIsErrorEverOccurred = true;
                            continue;
                        }



                        g_lsTestLogs = CollectTestLogs(sItem, g_sTestLogPath);

                        // update logger
                        Logger.Print($"{g_lsTestLogs.Count} log(s) are collected.", Logger.PrintOption.Both);
         


                        // Search and collect all TDS files.
                        // * Input: Start path
                        // * Output: A list containing the TDS file names.
                        g_lsTDSFiles = SearchTDSFiles(sTDSPath, sListFileForTDS);
                        if ((null == g_lsTDSFiles) || (0 == g_lsTDSFiles.Count))
                        {
                            continue;
                        }

                        Logger.Print($"{g_lsTDSFiles.Count} TDS are collected.", Logger.PrintOption.Both);



                        // Parse the function & TC info from each TDS file.
                        if (!ReadDataFromTDSFiles(sItem, ref g_lsTDSFiles, ref g_lsTestLogs))
                        {
                            bIsErrorEverOccurred = true;
                            continue;
                        }




                        // Count designs and methods.
                        CountAndMarkResults();



                        Logger.Print("Write the result to report.");

                        // Write the results to as an overall lookup table.
                        if (!SaveResults(g_sTemplateFile, sOutputFile, ref g_lsTestLogs))
                        {
                            bIsErrorEverOccurred = true;
                        }


                        //
                        // add iteminfo
                        //
                        ModuleInfo item = new ModuleInfo();

                        item.name = sItem;

                        item.gettersetter = gn_GetterSetter;
                        item.emptymethod = gn_Emptymethod;
                        item.abstractmethod = gn_Abstractmethod;
                        item.interfacemethod = gn_Interfacemethod;
                        item.nativemethod = gn_Nativemethod;

                        item.mockito = gn_ByMockito;
                        item.powermockito = gn_ByPowerMockito;

                        item.codeanalysis = gn_Bycodeanalysis;
                        item.purefunctioncalls = gn_Purefunctioncalls;
                        item.pureUIfunctioncalls = gn_PureUIfunctioncalls;

                        item.unknow = gn_Unknow;

                        item.count = g_tTestCaseTable.ltItems.Count;
                        item.testCase = g_tTestCaseTable;

                        //item.logNumError = logNumError;

                        // push item into a List
                        g_lsModuleInfo.Add(item);


                        Logger.Print("Writing result to Summary Reports.");

                        int index = GetModuleId(sItem, nameInSummary);

                        // write summary report
                        if (!WriteSummaryReport(sSummaryReportPath, item, index))
                        {

                        }

                    }
                    finally
                    {

                        #region Show the result when current module has checked.

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
                        Logger.Print("  Test means summary:", "");
                        Logger.Print("   - No test needed:         ", String.Format("{0,4}", g_tTestCaseTable.dByNACount.ToString()));
                        Logger.Print("     Getter/Setter:          ", String.Format("{0,4}", gn_GetterSetter));
                        Logger.Print("     Empty method:           ", String.Format("{0,4}", gn_Emptymethod));
                        Logger.Print("     Abstract method:        ", String.Format("{0,4}", gn_Abstractmethod));
                        Logger.Print("     Interface method:       ", String.Format("{0,4}", gn_Interfacemethod));
                        Logger.Print("     Native method:          ", String.Format("{0,4}", gn_Nativemethod));
                        Logger.Print("", "");

                        Logger.Print("   - Tested by test scripts: ", String.Format("{0,4}", g_tTestCaseTable.dByTestScriptCount.ToString()));
                        Logger.Print("     By Mockito:             ", String.Format("{0,4}", gn_ByMockito));
                        Logger.Print("     By PowerMockito:        ", String.Format("{0,4}", gn_ByPowerMockito));
                        Logger.Print("", "");

                        Logger.Print("   - Tested by code analysis:", String.Format("{0,4}", g_tTestCaseTable.dByCodeAnalysisCount.ToString()));
                        Logger.Print("     By code analysis:       ", String.Format("{0,4}", gn_Bycodeanalysis));
                        Logger.Print("     Pure function calls:    ", String.Format("{0,4}", gn_Purefunctioncalls));
                        Logger.Print("     Pure UI function calls: ", String.Format("{0,4}", gn_PureUIfunctioncalls));
                        Logger.Print("", "");

                        Logger.Print("   - Unknow items:           ", String.Format("{0,4}", g_tTestCaseTable.dByUnknownCount.ToString()));
                        Logger.Print("     Uknow:                  ", String.Format("{0,4}", gn_Unknow));
                        Logger.Print("", "");


                        dSum = g_tTestCaseTable.dByTestScriptCount + g_tTestCaseTable.dByCodeAnalysisCount +
                            g_tTestCaseTable.dByUnknownCount + g_tTestCaseTable.dByNACount;
                        if (dSum != g_tTestCaseTable.dNormalEntryCount)
                        {
                            Logger.Print("     (Error:", dSum.ToString() + " != " + g_tTestCaseTable.dNormalEntryCount.ToString() + ")");
                        }



                        // Show the overall NG entry #.
                        if (0 < g_tTestCaseTable.dNGEntryCount)
                        {
                            Logger.Print("  Total # of NG entry(s) found:", g_tTestCaseTable.dNGEntryCount.ToString());
                        }

                        #endregion

                        #region Accumulate the counts.

                        // Accumulate the counts.
                        dNormalEntryCount = dNormalEntryCount + g_tTestCaseTable.dNormalEntryCount;
                        dRepeatedEntryCount = dRepeatedEntryCount + g_tTestCaseTable.dRepeatedEntryCount;
                        dErrorEntryCount = dErrorEntryCount + g_tTestCaseTable.dErrorEntryCount;

                        dTestCaseFuncCount = dTestCaseFuncCount + g_tTestCaseTable.dTestCaseFuncCount;
                        dErrorCount = dErrorCount + g_tTestCaseTable.dErrorCount;
                        dNGEntryCount = dNGEntryCount + g_tTestCaseTable.dNGEntryCount;

                        dByMockito = dByMockito + gn_ByMockito;
                        dByPowerMockito = dByPowerMockito + gn_ByPowerMockito;
                        dBycodeanalysis = dBycodeanalysis + gn_Bycodeanalysis;
                        dGetterSetter = dGetterSetter + gn_GetterSetter;
                        dEmptymethod = dEmptymethod + gn_Emptymethod;
                        dAbstractmethod = dAbstractmethod + gn_Abstractmethod;
                        dInterfacemethod = dInterfacemethod + gn_Interfacemethod;
                        dNativemethod = dNativemethod + gn_Nativemethod;
                        dPurefunctioncalls = dPurefunctioncalls + gn_Purefunctioncalls;
                        dPureUIfunctioncalls = dPureUIfunctioncalls + gn_PureUIfunctioncalls;
                        dUnknow = dUnknow + gn_Unknow;

                        #endregion

                    }

                    // update progress
                    if (value < 90)
                    {
                        value = value + diff;
                        Logger.UpdateProgress(value);
                    }

                } // End of foreach

                // update logger
                Logger.UpdateProgress(90);

            }
            finally
            {


                // Show overall summary info.
                Logger.Print("", "---------------------------------------------------------------", Logger.PrintOption.Both);
                Logger.Print("", " Overall Summary:", Logger.PrintOption.Both);
                Logger.Print("", "---------------------------------------------------------------", Logger.PrintOption.Both);

                Logger.Print(" Total # of test cases with repeated labels: " + dRepeatedEntryCount.ToString(), "", Logger.PrintOption.Both);
                Logger.Print(" Total # of non-repeated test cases collected: " + dNormalEntryCount.ToString(), "", Logger.PrintOption.Both);
                Logger.Print(" Total # of non-repeated test case functions collected: " + dTestCaseFuncCount.ToString(), "", Logger.PrintOption.Both);
                Logger.Print(" Total # of errors found: " + dErrorCount.ToString(), "", Logger.PrintOption.Both);
                Logger.Print(" Total # of NG entries found: " + dNGEntryCount.ToString(), "", Logger.PrintOption.Both);
                Logger.Print(" Total # of Getter/Setter:          ", String.Format("{0,4}", dGetterSetter), Logger.PrintOption.Both);
                Logger.Print(" Total # of Empty method:           ", String.Format("{0,4}", dEmptymethod), Logger.PrintOption.Both);
                Logger.Print(" Total # of Abstract method:        ", String.Format("{0,4}", dAbstractmethod), Logger.PrintOption.Both);
                Logger.Print(" Total # of Interface method:       ", String.Format("{0,4}", dInterfacemethod), Logger.PrintOption.Both);
                Logger.Print(" Total # of Native method:          ", String.Format("{0,4}", dNativemethod), Logger.PrintOption.Both);
                Logger.Print(" Total # of By Mockito:             ", String.Format("{0,4}", dByMockito), Logger.PrintOption.Both);
                Logger.Print(" Total # of By PowerMockito:        ", String.Format("{0,4}", dByPowerMockito), Logger.PrintOption.Both);
                Logger.Print(" Total # of By code analysis:       ", String.Format("{0,4}", dBycodeanalysis), Logger.PrintOption.Both);
                Logger.Print(" Total # of Pure function calls:    ", String.Format("{0,4}", dPurefunctioncalls), Logger.PrintOption.Both);
                Logger.Print(" Total # of Pure UI function calls: ", String.Format("{0,4}", dPureUIfunctioncalls), Logger.PrintOption.Both);
                Logger.Print(" Total # of Uknow:                  ", String.Format("{0,4}", dUnknow), Logger.PrintOption.Both);


                // release Office.
                ReleaseOfficeApps();

                // Show ending message.
                if (bIsErrorEverOccurred)
                {
                    Logger.Print("Failed!");
                }
                else
                {

                    Logger.Print("Update path setting.");
                    WriteSetting();

                }
            }

            DateTime finishTime = DateTime.Now;

            // update looger
            Logger.Print("All Jobs Done! " + finishTime.ToString(new CultureInfo("en-US")), Logger.PrintOption.Both);

            TimeSpan timeSpan = finishTime.Subtract(startTime);
            Logger.Print($"Elapsed Time: {TimeSpan.ToString()}");

            Logger.UpdateProgress(100);

            return 0;
        }




        /// <summary>
        /// Init all variables
        /// </summary>
        /// <returns></returns>
        private bool InitializeVariable()
        {
            string sFuncName = "[Init]";

            // init the list of modules
            g_lsModules = new List<string>();

            // get the handler for excel
            g_excelApp = new Excel.Application
            {
                DisplayAlerts = false
            };

            // Initialize ...
            g_lsTDSFiles = new List<string>();

            // New the list for storing the data read.
            g_tTestCaseTable.ltItems = new List<TestCaseItem>();

            g_lsModuleInfo = new List<ModuleInfo>();
            g_lsModuleInfo.Clear();

            Logger.Print(sFuncName, "Done.");

            return true;
        }




        /// <summary>
        /// Update the setting of environment from Command to MainFrom, and Mainform to Globals
        /// </summary>
        /// <param name="commandline"></param>
        /// <returns></returns>
        public bool UpdateSetting()
        {
            string sFuncName = "[UpdatePathSetting]";

            EnvrionmentSetting ps = new EnvrionmentSetting();

            bool bDone = false;

            char[] delimiter = new char[] { ' ', '"' };
            string[] args = Environment.CommandLine.Split(delimiter, StringSplitOptions.RemoveEmptyEntries);



            if (args.Length == Constants.CommandArguments.Match)
            {

                // update setting to textbox
                ps.listFile = args[1];
                ps.tdsPath = args[2];
                ps.outputPath = args[3];
                ps.reportTemplate = args[4];
                ps.summaryTemplate = args[5];
                ps.testlogPath = args[6];

                g_FilePathSetting = ps;
                UpdatePathEvent(this, null);


                this.RunUTCheckerBy = RunBy.CommandLine;

                Logger.Print(sFuncName, "Update the path setting from Command Line.");

                bDone = true;

            }
            else
            {

                Logger.Print(sFuncName, "Invalid Arguments, Check setting file.");

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

                            }
                        }
                    }

                    g_FilePathSetting = ps;
                    UpdatePathEvent(this, null);

                    Logger.Print(sFuncName, "Update the path setting from UTChecker.setting");

                    bDone = true;
                }
                else
                {
                    Logger.Print(sFuncName, "The path settings haven't been updated.");
                    bDone = true;
                }

                this.RunUTCheckerBy = RunBy.User;
            }


            Logger.Print(sFuncName, "Done");

            return bDone;
        }



        /// <summary>
        /// Check the setting of enviroment before starting to check the UT 
        /// </summary>
        /// <returns></returns>
        private bool CheckSetting()
        {

            string sFuncName = "[CheckSetting]";

            EnvrionmentSetting fp = g_MF.GetPath();

            g_FilePathSetting = fp;

            g_sModuleListFile =  fp.listFile;           // list file
            g_sTDSPath = fp.tdsPath;                    // tds path
            g_sOutputPath = fp.outputPath;              // output path
            g_sTemplateFile = fp.reportTemplate; ; ;    // template file
            g_sSummaryReport = fp.summaryTemplate;      // summary templat
            g_sTestLogPath = fp.testlogPath;

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
                Logger.Print(sFuncName, "Cannot find the TDS path: " + g_sTDSPath);
                return false;
            }
            if (!Directory.Exists(g_sOutputPath))
            {
                Logger.Print(sFuncName, "Cannot find the output path: " + g_sOutputPath);
                return false;
            }

            if (!Directory.Exists(g_sTestLogPath))
            {
                Logger.Print(sFuncName, "Cannot find the output path: " + g_sTestLogPath);
                return false;
            }


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

            return true;

        }



        /// <summary>
        /// Write the setting of environment to file.
        /// </summary>
        /// <returns></returns>
        private bool WriteSetting()
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

            Logger.Print(sFuncName, "Done.");

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
                    }
                }

                // Record the number of non-repeated test cases.
                g_tTestCaseTable.dNormalEntryCount = dNormalEntryCount;
                g_tTestCaseTable.dRepeatedEntryCount = dRepeatedEntryCount;
                g_tTestCaseTable.dErrorEntryCount = dErrorEntryCount;

                g_tTestCaseTable.dByNACount = dDoneByNACount;
                g_tTestCaseTable.dByTestScriptCount = dDoneByTestScriptCount;
                g_tTestCaseTable.dByCodeAnalysisCount = dDoneByCodeAnalysisCount;
                g_tTestCaseTable.dByUnknownCount = dDoneByOthersCount;

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
