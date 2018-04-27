using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace UTChecker
{
    public partial class TDS_Parser
    {
        private MainForm g_MF = null;

        private LoggerForm g_LF = null;



        /// <summary>
        /// Constructor
        /// </summary>
        public TDS_Parser(MainForm mf, LoggerForm lf)
        {

            g_MF = mf;
            g_LF = lf;


            g_FilePathSetting = new PathSetting();

            // init the log file for recoding the process.
            InitializeProcessLog();

            // init background
            //InitializeBackgroundWorkerForLogToWindow();

            InitializeBackgroundWorkerForTDSParse();
        }


        /// <summary>
        /// Init a backgroundworker for log message to listbox
        /// </summary>
        public void InitializeBackgroundWorkerForTDSParse()
        {
            g_bwTDSParse = new BackgroundWorker();

            g_bwTDSParse.WorkerReportsProgress = true;
            g_bwTDSParse.WorkerSupportsCancellation = true;
            g_bwTDSParse.DoWork += new DoWorkEventHandler(bwTDSParse_DoWork);
            g_bwTDSParse.ProgressChanged += new ProgressChangedEventHandler(bwTDSParse_ProgressChanged);
            g_bwTDSParse.RunWorkerCompleted += new RunWorkerCompletedEventHandler(bwTDSParse_RunWorkerCompleted);
        }

        // event for DoWork
        private void bwTDSParse_DoWork(object sender, DoWorkEventArgs e)
        {

            // Get the BackgroundWorker that raised this event.
            BackgroundWorker worker = sender as BackgroundWorker;
            e.Result = RunTDSParser();
        }

        // event for ProgressChanged
        private void bwTDSParse_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            //this.progressBarUTChecker.Value = e.ProgressPercentage;
        }

        // event for 
        private void bwTDSParse_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (this.RunUTCheckerBy == RunBy.CommandLine)
            {
                Environment.ExitCode = 0;
                Environment.Exit(Environment.ExitCode);
            }
            else
            {
                MessageBox.Show("Done");
                //this.progressBarUTChecker.Value = 0;
            }
        }


        /// <summary>
        /// 
        /// </summary>
        public void Run()
        {
            g_bwTDSParse.RunWorkerAsync();
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

            LogToFile(sFuncName, "Done.");

            return true;
        }


        /// <summary>
        /// Main routine for TDS_Parse
        /// </summary>
        /// <returns></returns>
        public int RunTDSParser()
        {
            string sFuncName = "[TDS_Parser]";


            string sStartPath;
            string sListFile;
            string sOutputFile;

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

            ClearMessage();
            ClearProgress();

            LogToWindow("Run TDS Parser.");


            // initial all variables
            InitializeVariable();


            ReportProgress(10);


            if (!CheckSetting())
            {
                Environment.ExitCode = -1;
                Environment.Exit(Environment.ExitCode);
            }

            LogToWindow("Initialize variables done.");


            ReportProgress(20);


            // Read the module list, where comment/empty lines will be ignored.
            if (!ReadTextFileToStringList(g_sModuleListFile, ref g_lsModules, true, true))
            {
                LogToFile(sFuncName, "Read DD module list failed.");
                LogToWindow("Read DD module list failed.");
                return 0;
            }


            ReportProgress(30);


            // prepare summary report
            string sSummaryReport = PrepareSummaryReport(g_sOutputPath);

            int diff = 40 / g_lsModules.Count;
            int value = 40;

            LogToWindow($"Total {g_lsModules.Count} modules would be checked.");

            try
            {

                foreach (string sItem in g_lsModules)
                {
                    // Determine the input/output file names.
                    sStartPath = g_sTDSPath + sItem;
                    sListFile = sStartPath + "_TDS.list";
                    sOutputFile = g_sOutputPath + TestCaseTableConstants.FILENAME_PREFIX + sItem + ".xlsx";
                    g_sErrorLogFile = g_sOutputPath + TestCaseTableConstants.FILENAME_PREFIX + sItem + ".log";


                    LogToWindow($"{sItem} is processing now.");

                    // Remove old log file.
                    if (File.Exists(g_sErrorLogFile))
                    {
                        File.Delete(g_sErrorLogFile);
                    }


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

                    // Write the spliter for reading the error log file easily.
                    // (This section must be located behind the remove-error-file section. Otherwise the message will be written to the previous error log file.)
                    LogToFile("", "---------------------------------------------------------------");
                    LogToFile("", sItem);
                    LogToFile("", "---------------------------------------------------------------");


                    try
                    {

                        // Prepare a dummy output report in case the flow is aborted due to a failure.
                        if (!PrepareDummayOutputReport(sOutputFile))
                        {
                            bIsErrorEverOccurred = true;
                            continue;
                        }

                        // Search and collect all TDS files.
                        // * Input: Start path
                        // * Output: A list containing the TDS file names.
                        SearchTDSFiles(sStartPath, ref g_lsTDSFiles, sListFile);
                        if (0 == g_lsTDSFiles.Count)
                        {
                            continue;
                        }

                        // Parse the function & TC info from each TDS file.
                        if (!ReadDataFromTDSFiles(sItem, ref g_lsTDSFiles))
                        {
                            bIsErrorEverOccurred = true;
                            continue;
                        }


                        // Count designs and methods.
                        CountAndMarkResults();

                        LogToWindow("Writing result to excel.");

                        // Write the results to as an overall lookup table.
                        if (!SaveResults(g_sTemplateFile, sOutputFile))
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

                        // push item into a List
                        g_lsModuleInfo.Add(item);

                    }
                    finally
                    {

                        // Show processed results.

                        // Show the # based on TDS entries. 
                        LogToFile("  Total # of test cases defined in TDS:", g_tTestCaseTable.ltItems.Count.ToString());
                        LogToFile("   - Normal entries:  ", String.Format("{0,4}", g_tTestCaseTable.dNormalEntryCount.ToString()));
                        LogToFile("   - Repeated entries:", String.Format("{0,4}", g_tTestCaseTable.dRepeatedEntryCount.ToString()));
                        LogToFile("   - Error entries:   ", String.Format("{0,4}", g_tTestCaseTable.dErrorEntryCount.ToString()));
                        int dSum = g_tTestCaseTable.dNormalEntryCount + g_tTestCaseTable.dRepeatedEntryCount + g_tTestCaseTable.dErrorEntryCount;
                        if (dSum != g_tTestCaseTable.ltItems.Count)
                        {
                            LogToFile("     (Error:", dSum.ToString() + " != " + g_tTestCaseTable.ltItems.Count.ToString() + ")");
                        }

                        // new line
                        LogToFile("", "");

                        // Show the # based on test means. 
                        LogToFile("  Test means summary:", "");
                        LogToFile("   - No test needed:         ", String.Format("{0,4}", g_tTestCaseTable.dByNACount.ToString()));
                        LogToFile("     Getter/Setter:          ", String.Format("{0,4}", gn_GetterSetter));
                        LogToFile("     Empty method:           ", String.Format("{0,4}", gn_Emptymethod));
                        LogToFile("     Abstract method:        ", String.Format("{0,4}", gn_Abstractmethod));
                        LogToFile("     Interface method:       ", String.Format("{0,4}", gn_Interfacemethod));
                        LogToFile("     Native method:          ", String.Format("{0,4}", gn_Nativemethod));
                        LogToFile("", "");

                        LogToFile("   - Tested by test scripts: ", String.Format("{0,4}", g_tTestCaseTable.dByTestScriptCount.ToString()));
                        LogToFile("     By Mockito:             ", String.Format("{0,4}", gn_ByMockito));
                        LogToFile("     By PowerMockito:        ", String.Format("{0,4}", gn_ByPowerMockito));
                        LogToFile("", "");

                        LogToFile("   - Tested by code analysis:", String.Format("{0,4}", g_tTestCaseTable.dByCodeAnalysisCount.ToString()));
                        LogToFile("     By code analysis:       ", String.Format("{0,4}", gn_Bycodeanalysis));
                        LogToFile("     Pure function calls:    ", String.Format("{0,4}", gn_Purefunctioncalls));
                        LogToFile("     Pure UI function calls: ", String.Format("{0,4}", gn_PureUIfunctioncalls));
                        LogToFile("", "");

                        LogToFile("   - Unknow items:           ", String.Format("{0,4}", g_tTestCaseTable.dByUnknownCount.ToString()));
                        LogToFile("     Uknow:                  ", String.Format("{0,4}", gn_Unknow));
                        LogToFile("", "");


                        dSum = g_tTestCaseTable.dByTestScriptCount + g_tTestCaseTable.dByCodeAnalysisCount +
                            g_tTestCaseTable.dByUnknownCount + g_tTestCaseTable.dByNACount;
                        if (dSum != g_tTestCaseTable.dNormalEntryCount)
                        {
                            LogToFile("     (Error:", dSum.ToString() + " != " + g_tTestCaseTable.dNormalEntryCount.ToString() + ")");
                        }



                        // Show the overall NG entry #.
                        if (0 < g_tTestCaseTable.dNGEntryCount)
                        {
                            LogToFile("  Total # of NG entry(s) found:", g_tTestCaseTable.dNGEntryCount.ToString());
                        }

#if !DEBUG
                        Console.WriteLine("For details, please check the log file (" + Path.GetFileName(g_sErrorLogFile) + ").\n");
#endif

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


                    }


                    if (value < 80)
                    {
                        value = value + diff;
                        ReportProgress(value);
                    }

                    


                } // End of foreach

                ReportProgress(80);

            }
            finally
            {
                LogToWindow("Writing Summary Reports.");
                ReportProgress(90);


                // write summary report
                if (!WriteSummaryReport(sSummaryReport, ref g_lsModuleInfo))
                {

                }



                // Show overall summary info.
                LogToFileAndWin("", "\n---------------------------------------------------------------");
                LogToFileAndWin("", " Overall Summary:");
                LogToFileAndWin("", "---------------------------------------------------------------");

                LogToFileAndWin(" Total # of test cases with repeated labels: " + dRepeatedEntryCount.ToString(), "");
                LogToFileAndWin(" Total # of non-repeated test cases collected: " + dNormalEntryCount.ToString(), "");
                LogToFileAndWin(" Total # of non-repeated test case functions collected: " + dTestCaseFuncCount.ToString(), "");
                LogToFileAndWin(" Total # of errors found: " + dErrorCount.ToString(), "");
                LogToFileAndWin(" Total # of NG entries found: " + dNGEntryCount.ToString(), "");
                LogToFileAndWin(" Total # of Getter/Setter:          ", String.Format("{0,4}", dGetterSetter));
                LogToFileAndWin(" Total # of Empty method:           ", String.Format("{0,4}", dEmptymethod));
                LogToFileAndWin(" Total # of Abstract method:        ", String.Format("{0,4}", dAbstractmethod));
                LogToFileAndWin(" Total # of Interface method:       ", String.Format("{0,4}", dInterfacemethod));
                LogToFileAndWin(" Total # of Native method:          ", String.Format("{0,4}", dNativemethod));
                LogToFileAndWin(" Total # of By Mockito:             ", String.Format("{0,4}", dByMockito));
                LogToFileAndWin(" Total # of By PowerMockito:        ", String.Format("{0,4}", dByPowerMockito));
                LogToFileAndWin(" Total # of By code analysis:       ", String.Format("{0,4}", dBycodeanalysis));
                LogToFileAndWin(" Total # of Pure function calls:    ", String.Format("{0,4}", dPurefunctioncalls));
                LogToFileAndWin(" Total # of Pure UI function calls: ", String.Format("{0,4}", dPureUIfunctioncalls));
                LogToFileAndWin(" Total # of Uknow:                  ", String.Format("{0,4}", dUnknow));



                ReleaseOfficeApps();



                // Show ending message.
                if (bIsErrorEverOccurred)
                {
                    LogToWindow("Failed!");
                }
                else
                {

                    LogToWindow("Update path setting.");
                    UpdateUTCheckerSettingToFile();

                    LogToWindow("All Jobs Done!");



                }
            }

            ReportProgress(100);


            return 0;
        }



        /// <summary>
        /// parse the path of setting.
        /// </summary>
        /// <param name="commandline"></param>
        /// <returns></returns>
        public bool UpdatePathSetting()
        {
            string sFuncName = "[UpdatePathSetting]";

            PathSetting ps = new PathSetting();

            bool bDone = false;

            char[] delimiter = new char[] { ' ', '"' };
            string[] args = Environment.CommandLine.Split(delimiter, StringSplitOptions.RemoveEmptyEntries);

            if (args.Length == Constants.ArgumentsMatchLength)
            {

                // update setting to textbox
                ps.listFile = args[1];
                ps.tdsPath = args[2];
                ps.outputPath = args[3];
                ps.reportTemplate = args[4];
                ps.summaryTemplate = args[5];

                g_FilePathSetting = ps;
                UpdatePathEvent(this, null);


                this.RunUTCheckerBy = RunBy.CommandLine;

                LogToFile(sFuncName, "Update the path setting from Command Line.");

                bDone = true;

            }
            else
            {
                if ((args.Length > Constants.ArgumentsMatchLength) ||
                    ((args.Length < Constants.ArgumentsMatchLength) &&
                    args.Length > Constants.UTCheckerSelf))
                {
                    LogToFile(sFuncName, "Invalid Arguments.");
                    bDone = false;
                }
                else if (File.Exists(UTCheckerSetting.FileName))
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
                                else if (setting[1].Equals(UTCheckerSetting.TestLogs))
                                {
                                    ps.testlogsPath = setting[2];
                                }

                            }
                        }
                    }

                    g_FilePathSetting = ps;
                    UpdatePathEvent(this, null);

                    LogToFile(sFuncName, "Update the path setting from UTChecker.setting");

                    bDone = true;
                }
                else
                {
                    LogToFile(sFuncName, "The path settings haven't been updated.");
                    bDone = true;
                }

                this.RunUTCheckerBy = RunBy.User;
            }


            LogToFile(sFuncName, "Done");

            return bDone;
        }



        /// <summary>
        /// Check the setting of path before starting to check the UT 
        /// </summary>
        /// <returns></returns>
        private bool CheckSetting()
        {

            string sFuncName = "[CheckSetting]";

            PathSetting fp = g_MF.GetPath();

            g_FilePathSetting = fp;

            g_sModuleListFile =  fp.listFile;            // list file
            g_sTDSPath = fp.tdsPath;                    // tds path
            g_sOutputPath = fp.outputPath;              // output path
            g_sTemplateFile = fp.reportTemplate; ; ;    // template file
            g_sSummaryReport = fp.summaryTemplate;      // summary templat


            // Ensure each path is ended with a '\\'.
            if (!g_sTDSPath.EndsWith("\\"))
            {
                g_sTDSPath = g_sTDSPath + "\\";
            }

            if (!g_sOutputPath.EndsWith("\\"))
            {
                g_sOutputPath = g_sOutputPath + "\\";
            }


            // Check the existence of the input & output paths.
            if (!Directory.Exists(g_sTDSPath))
            {
                LogToFile(sFuncName, "Cannot find the TDS path: " + g_sTDSPath);
                return false;
            }
            if (!Directory.Exists(g_sOutputPath))
            {
                LogToFile(sFuncName, "Cannot find the output path: " + g_sOutputPath);
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
                LogToFile(sFuncName, "Cannot find Module List File: " + g_sModuleListFile);
                return false;
            }
            if (!File.Exists(g_sTemplateFile))
            {
                LogToFile(sFuncName, "Cannot find the template file: " + g_sTemplateFile);
                return false;
            }

            if (!File.Exists(g_sSummaryReport))
            {
                LogToFile(sFuncName, "Cannot find the summary report: " + g_sSummaryReport);
               return false;
            }

            return true;

        }


        /// <summary>
        /// 
        /// </summary>
        /// <param name="a_sInFile"></param>
        /// <param name="a_lsOutList"></param>
        /// <param name="a_bTrim"></param>
        /// <param name="a_bValidLinesOnly"></param>
        /// <returns></returns>
        private bool ReadTextFileToStringList(string a_sInFile, ref List<string> a_lsOutList, bool a_bTrim, bool a_bValidLinesOnly)
        {
            string sFuncName = "[ReadTextFileToStringList]";

            // Check the existence of the input file.
            if (!File.Exists(a_sInFile))
            {
                LogToFile(sFuncName, "Cannot find \"" + a_sInFile + "\"");
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
                            a_lsOutList[i] = a_lsOutList[i].Trim();
                    }

                    // Remove invalid lines from the list:
                    // (1) Comment lines (line starts with a '#')
                    // (2) Empty lines
                    // (3) Lines contain spaces only
                    if (a_bValidLinesOnly)
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
                                a_lsOutList.RemoveAt(i);
                        }
                }

                if (0 == a_lsOutList.Count)
                {
                    LogToFile(sFuncName, "No-line is loaded from " + a_sInFile);
                }
            }
            catch (Exception ex)
            {
                LogToFile(sFuncName, ex.ToString());
                return false;
            }

            LogToFile(sFuncName, "Done");

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
                    LogToFile(sFuncName, "Cannot find " + g_sTemplateFile);
                    return false;
                }
                File.Copy(g_sTemplateFile, a_sOutputFile);
            }
            catch (Exception e)
            {
                LogToFile(sFuncName, "Exception: " + e.ToString());
                return false;
            }

            return true;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="a_sStartPath"></param>
        /// <param name="a_lsOutList"></param>
        /// <param name="a_sOutFile"></param>
        /// <returns></returns>
        private bool SearchTDSFiles(string a_sStartPath, ref List<string> a_lsOutList, string a_sOutFile)
        {
            string sFuncName = "[SearchTDSFiles]";

            // Check the input parameters.
            if ("" == a_sStartPath)
            {
                LogToFile(sFuncName, "Null start path is specified.");
                return false;
            }

            // New a list if needs.
            if (null == a_lsOutList)
            {
                a_lsOutList = new List<string>();
            }
            else
            {
                a_lsOutList.Clear();
            }

            // Serach and collect all log files recursively.
            CollectFiles(a_sStartPath, TestCaseTableConstants.INPUT_FILE_EXT_NAME, TestCaseTableConstants.INPUT_FILENAME_PREFIX, ref a_lsOutList);

            // Save the list of found files to the specifed file.
            if ("" != a_sOutFile)
            {
                WriteStringListToTextFile(ref a_lsOutList, a_sOutFile);
            }

            LogToFile(sFuncName, a_lsOutList.Count.ToString() + " TDS file(s) collected.");

            return true;
        }


        /// <summary>
        /// 
        /// </summary>
        /// <param name="a_sDir"></param>
        /// <param name="a_sFileExt"></param>
        /// <param name="a_sToken"></param>
        /// <param name="a_lsCollection"></param>
        /// <returns></returns>
        private List<string> CollectFiles(string a_sDir, string a_sFileExt, string a_sToken, ref List<string> a_lsCollection)
        {
            string sFuncName = "[CollectFiles]";
            string sFileName;

            try
            {
                // Check the existence of the specified path.
                if (!Directory.Exists(a_sDir))
                {
                    LogToFile(sFuncName, "Cannot find path \"" + a_sDir + "\"; skipped.");
                    return a_lsCollection;
                }

                // Collect the considered files stored in current folder.
                string[] FileList = Directory.GetFiles(a_sDir, a_sFileExt);
                foreach (string f in FileList)
                {
                    // Discard the path from the name.
                    sFileName = Path.GetFileName(f);

                    // Check if the file name starts with the spcified token.
                    // If yes, add it in the list.
                    if (sFileName.StartsWith(a_sToken))
                        a_lsCollection.Add(f);
                }

                // Collect the considered files stored in sub-folders.
                string[] DirList = Directory.GetDirectories(a_sDir);
                foreach (string d in DirList)
                {
                    a_lsCollection = CollectFiles(d, a_sFileExt, a_sToken, ref a_lsCollection);
                }
            }
            catch (System.Exception excpt)
            {
                LogToFile(sFuncName, excpt.Message);
            }

            return a_lsCollection;
        }


        /// <summary>
        /// 
        /// </summary>
        /// <param name="a_lsInList"></param>
        /// <param name="a_sOutFile"></param>
        /// <returns></returns>
        private bool WriteStringListToTextFile(ref List<string> a_lsInList, string a_sOutFile)
        {
            string sFuncName = "[WriteStringListToTextFile]";

            // Check the input.
            if (null == a_lsInList)
            {
                LogToFile(sFuncName, "Cannot save a null list to file.");
                return false;
            }
            if ("" == a_sOutFile)
            {
                LogToFile(sFuncName, "No output file is specified.");
                return false;
            }

            // Check the number of lines to be saved.
            if (0 == a_lsInList.Count)
            {
                LogToWindow(sFuncName + "The list to be saved is an empty list. Do nothing.");
                return true;
            }

            // Write the error log to the output file.
            try
            {
                using (StreamWriter sw = File.CreateText(a_sOutFile))
                {
                    foreach (string sLine in a_lsInList)
                    {
                        sw.WriteLine(sLine);
                    }
                }
            }
            catch (Exception ex)
            {
                LogToFile(sFuncName, ex.ToString());
                return false;
            }

            return true;
        }


        /// <summary>
        /// Determine the type/mean for test case.
        /// </summary>
        /// <param name="a_sInfo"></param>
        /// <returns></returns>
        private TestMeans DetermineTestMeans(string a_sInfo)
        {
            TestMeans eTestMeans = TestMeans.UNKNOWN;

            if (a_sInfo.Equals("N/A"))
            {
                eTestMeans = TestMeans.TEST_SCRIPT;
                gn_ByMockito++;
            }
            else if (a_sInfo.Equals(TestType.ByPowerMocktio))
            {
                eTestMeans = TestMeans.TEST_SCRIPT;
                gn_ByPowerMockito++;
            }
            else if (a_sInfo.Equals(TestType.ByCodeAnalysis))
            {
                eTestMeans = TestMeans.CODE_ANALYSIS;
                gn_Bycodeanalysis++;
            }
            else if (a_sInfo.Equals(TestType.GetterSetter))
            {
                eTestMeans = TestMeans.NA;
                gn_GetterSetter++;
            }
            else if (a_sInfo.Equals(TestType.Empty))
            {
                eTestMeans = TestMeans.NA;
                gn_Emptymethod++;
            }
            else if (a_sInfo.Equals(TestType.Abstract))
            {
                eTestMeans = TestMeans.NA;
                gn_Abstractmethod++;
            }
            else if (a_sInfo.Equals(TestType.Interface))
            {
                eTestMeans = TestMeans.NA;
                gn_Interfacemethod++;
            }
            else if (a_sInfo.Equals(TestType.Native))
            {
                eTestMeans = TestMeans.NA;
                gn_Nativemethod++;
            }
            else if (a_sInfo.Equals(TestType.PureFunctionCalls))
            {
                //eMethodType = MethodType.PURE_CALL;
                eTestMeans = TestMeans.CODE_ANALYSIS;
                gn_Purefunctioncalls++;
            }
            else if (a_sInfo.Equals(TestType.PureUIFunctionCalss))
            {
                eTestMeans = TestMeans.CODE_ANALYSIS;
                gn_PureUIfunctioncalls++;
            }
            else
            {
                eTestMeans = TestMeans.UNKNOWN;
                gn_Unknow++;

                LogToFile(" - UNKNOW: ", String.Format("\"{0}\"", a_sInfo));
            }

            return eTestMeans;
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
                        ltNonRepeatedItems.RemoveAt(i);
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
                        ltNonRepeatedItems.RemoveAt(i);
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
                        ltNonRepeatedItems.RemoveAt(i);
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
                            dDoneByNACount++;
                        else if (TestMeans.TEST_SCRIPT == tTestCase.eTestMeans)
                            dDoneByTestScriptCount++;
                        else if (TestMeans.CODE_ANALYSIS == tTestCase.eTestMeans)
                            dDoneByCodeAnalysisCount++;
                        else
                            dDoneByOthersCount++;
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
                    LogToFile("     (Error:", dSum.ToString() + " != " + g_tTestCaseTable.ltItems.Count.ToString() + ")");

                dSum = g_tTestCaseTable.dByNACount + g_tTestCaseTable.dByTestScriptCount +
                        g_tTestCaseTable.dByCodeAnalysisCount + g_tTestCaseTable.dByUnknownCount;
                if (dSum != g_tTestCaseTable.dNormalEntryCount)
                    LogToFile("     (Error:", dSum.ToString() + " != " + g_tTestCaseTable.dNormalEntryCount.ToString() + ")");
            }
            catch (SystemException e)
            {
                LogToFile(sFuncName, e.ToString());
            }
        }





    }
}
