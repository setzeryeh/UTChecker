using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;

namespace UTChecker
{
    public partial class UTChecker
    {

        #region Parse TDS


        /// <summary>
        /// Read the information from Coversheet.
        /// </summary>
        /// <param name="a_excelBook"></param>
        /// <param name="a_sSourceFileName"></param>
        /// <param name="a_sMethodName"></param>
        /// <returns></returns>
        private bool ReadInfoFromTDSCoverSheet(Excel.Workbook a_excelBook, ref string a_sSourceFileName, ref string a_sMethodName)
        {
            Excel.Worksheet excelSheet;
            Excel.Range excelRange;
            string sSourceFileName;
            string sValue;
            int dRow, dCol;
            const string sSheetName = "Coversheet";
            bool bIsJava, bIsCOrCpp;


            // -------------------------------------------------------------------------
            // Read data form the "Coversheet" sheet.
            // -------------------------------------------------------------------------
            try
            {
                excelSheet = (Excel.Worksheet)a_excelBook.Worksheets.get_Item(sSheetName);
                excelRange = excelSheet.UsedRange;
            }
            catch
            {
                Logger.Print(Constants.StringTokens.MSG_BULLET, "No \"" + sSheetName + "\" sheet can be found.");
                return false;
            }

            // Locate the considered cell (where y in [2,4] and x in [2,3]).
            bool bFound = false;
            dCol = 2;
            for (dRow = 2; dRow <= 3; dRow++)
            {
                for (dCol = 2; dCol <= 4; dCol++)
                {
                    sValue = ReadStringFromExcelCell(excelRange.Cells[dRow, dCol], "", true);
                    if ("File" == sValue)
                    {
                        bFound = true;
                        break;
                    }
                }
                if (bFound)
                {
                    break;
                }
            }

            if (bFound)
            {
                // Read the file name & SVN revision.
                sSourceFileName = ReadStringFromExcelCell(excelRange.Cells[dRow, dCol + 1], "", true);

                // Check if the source file name contains any space, and remove it.
                // e.g. "XXX .java" --> "XXX.java"
                if (sSourceFileName.Contains(" "))
                {
                    Logger.Print(Constants.StringTokens.MSG_BULLET, "File name \"" + sSourceFileName + "\" contains space(s). Stripped.");
                    sSourceFileName.Replace(" ", "");
                }

                // Check if the string read is a java/c/cpp file.
                string sTmp = sSourceFileName.ToLower();
                bIsJava = sTmp.EndsWith(".java");
                bIsCOrCpp = (sTmp.EndsWith(".c") || sTmp.EndsWith(".cpp"));

                // Set the source file name read and increase the row.
                if (bIsCOrCpp || bIsJava)
                {
                    a_sSourceFileName = sSourceFileName;
                    dRow++;
                }
            }
            else
            {
                Logger.Print(Constants.StringTokens.MSG_BULLET, "No source file name can be found from the \"" + sSheetName + "\" sheet.");
                return false;
            }

            // Extrace method name.
            a_sMethodName = "";
            if (bIsCOrCpp)
            {
                // Locate the method name cell.
                bFound = false;
                dRow++;
                for (; dRow <= 7; dRow++)
                {
                    for (dCol = 1; dCol <= 3; dCol++)
                    {
                        sValue = ReadStringFromExcelCell(excelRange.Cells[dRow, dCol], "", true);
                        if ("Method Name" == sValue)
                        {
                            bFound = true;
                            break;
                        }
                    }

                    // Extract method name.
                    if (bFound)
                    {
                        dRow++;
                        dCol = dCol + 2;
                        sValue = ReadStringFromExcelCell(excelRange.Cells[dRow, dCol], "", true);
                        a_sMethodName = ExtractMethodName(sValue);

                        return ("" != a_sMethodName);
                    }
                }

                return false;
            }

            return true;
        }


        /// <summary>
        /// 
        /// </summary>
        /// <param name="a_sModuleName"></param>
        /// <param name="a_lsTDSFiles"></param>
        /// <param name="a_lsTestLogs"></param>
        /// <param name="a_sutsDocPath"></param>
        /// <returns></returns>
        private bool ReadDataFromTDSFiles(string a_sModuleName, ref List<string> a_lsTDSFiles, ref List<TestLog> a_lsTestLogs, string a_SUTSDocPath)
        {
            string sFuncName = "[ReadDataFromTDSFiles]";

            Excel.Workbook excelBook = null;
            Word.Document wordDoc = null;

            bool bIsJava;
            string sFileNameWithoutPath;
            string sShortTDSFileName;
            string sSourceFileName = "";
            string sMethodName = "";
            int dErrorCount = 1;
            int dProceedFileCount = 0;


            // Check the input list.
            if (null == a_lsTDSFiles)
            {
                Logger.Print(sFuncName, "The input list of TDS is null.");
                return false;
            }

            if (0 == a_lsTDSFiles.Count)
            {
                Logger.Print(sFuncName, "The TDS file(s) aren't found.");
                return false;
            }

            // Check the input list.
            if (null == a_lsTestLogs)
            {
                Logger.Print(sFuncName, "The input list of Test log is null.");
                return false;
            }

            if (0 == a_lsTestLogs.Count)
            {
                Logger.Print(sFuncName, "The Test Log file(s) aren't found.");
                return false;
            }


            // Check the EXCEL app.
            if (null == g_excelApp)
            {
                Logger.Print(sFuncName, ErrorMessage.EXCEL_APP_IS_NULL);
                return false;
            }

            if (null == g_wordApp)
            {
                Logger.Print(sFuncName, ErrorMessage.WORD_APP_IS_NULL);
                return false;
            }


            try
            {
                Logger.Print(sFuncName, "Reading TDS files...");

                // Initialize objects.
                dErrorCount = 0;
                g_tTestCaseTable.ltItems.Clear();
                g_excelApp.DisplayAlerts = false; // show no alert while closing the file




                if (a_SUTSDocPath != String.Empty)
                {
                    wordDoc = OpenWordDocument(g_wordApp, a_SUTSDocPath);
                    if (null == wordDoc)
                    {
                        Logger.Print(sFuncName, $"Error occurred when Open SUTS.");
                    }

                    // clear previous result.
                    SUTS_ClearPreviousResult();

                }
                else
                {
                    Logger.Print(sFuncName, $"No SUTS, Skipped!");
                }


                // Read data from each TDS file.
                foreach (string sFile in a_lsTDSFiles)
                {
                    sFileNameWithoutPath = "\"" + Path.GetFileName(sFile) + "\"";
                    Logger.Print(sFuncName, "Reading " + sFileNameWithoutPath + "...");

                    // Check the existence of the TDS file.
                    if (!File.Exists(sFile))
                    {
                        Logger.Print(Constants.StringTokens.MSG_BULLET, "Cannot find " + sFile);
                        dErrorCount++;
                        continue;
                    }

                    // Open the TDS file & get the lookup-table sheet.
                    excelBook = OpenExcelWorkbook(g_excelApp, sFile, true);  // ready only
                    if (null == excelBook)
                    {
                        continue;
                    }

                   


                    try
                    {
                        // Read source file name from the cover sheet.
                        if (!ReadInfoFromTDSCoverSheet(excelBook, ref sSourceFileName, ref sMethodName))
                        {
                            dErrorCount++;
                            continue;
                        }


                        // Determine the source file type.
                        bIsJava = (sSourceFileName.EndsWith("java"));


                        // Read data form the "TestCase" sheet.
                        sShortTDSFileName = sFile.Replace(g_sTDSPath + a_sModuleName + "\\", "");
                        if (bIsJava)
                        {
                            ReadTestCasesFromTDSFile_Java(excelBook, wordDoc, ref sShortTDSFileName, ref sSourceFileName, ref a_lsTestLogs);
                        }
                        else
                        {
                            ReadTestCasesFromTDSFile_C(excelBook, wordDoc, ref sShortTDSFileName, ref sSourceFileName, ref sMethodName, ref a_lsTestLogs);
                        }


                        dProceedFileCount++;

                    }
                    catch (SystemException ex)
                    {
                        Logger.Print(sFuncName, sFile.Replace(g_sTDSPath, "...") + ": " + ex.ToString());
                        dErrorCount++;
                    }
                    finally
                    {
                        if (excelBook != null)
                        {
                            // Close the TDS file.
                            excelBook.Close(false, Type.Missing, Type.Missing);
                        }
                    }
                }


                if (wordDoc != null)
                {
                    wordDoc.Close((Object)Word.WdSaveOptions.wdDoNotSaveChanges, Type.Missing, Type.Missing);
                }


            }
            catch (SystemException ex)
            {
                Logger.Print(sFuncName, ex.ToString());
                dErrorCount++;
            }

            // Show the # of proceeded files.
            if (dProceedFileCount != a_lsTDSFiles.Count)
            {
                Logger.Print(sFuncName, dProceedFileCount.ToString() + " of " + a_lsTDSFiles.Count + " TDS files proceed.");
            }

            return true;
        }


        #endregion






        #region Java Group



        /// <summary>
        /// 
        /// </summary>
        /// <param name="a_sOrgMethodName"></param>
        /// <returns></returns>
        private string ArrangeAndCheckMethodName(string a_sOrgMethodName)
        {
            string sMethodName = "";
            string[] asElements;

            // Remove dummy "()".
            a_sOrgMethodName = a_sOrgMethodName.Replace("()", "");

            if (a_sOrgMethodName.Contains("(")) // e.g. XXX(int, int)
            {
                // Split the string.
                asElements = a_sOrgMethodName.Split('(');

                // Left part (method name) cannot contain any space
                if (asElements[0].Trim().Contains(" "))
                    return "";

                // Arrange the method name, e.g.
                // Case 1: XXX(int, int) --> XXX(int,int)
                // Case 2: XXX(  ) --> XXX()
                sMethodName = a_sOrgMethodName.Replace(" ", "");
                // Case: XXX() --> XXX
                sMethodName = sMethodName.Replace("()", "");

                return sMethodName;
            }
            else
            {
                if (a_sOrgMethodName.Contains(" ")) // method name cannot contain any space
                {
                    return "";
                }
                else
                {
                    return a_sOrgMethodName;
                }
            }
        }


        /// <summary>
        /// 
        /// </summary>
        /// <param name="a_excelBook"></param>
        /// <param name="a_sTDSFile"></param>
        /// <param name="a_sSourceFileName"></param>
        /// <returns></returns>
        private int ReadTestCasesFromTDSFile_Java(
                                    Excel.Workbook a_excelBook, 
                                    Word.Document a_wordDoc,
                                    ref string a_sTDSFile, 
                                    ref string a_sSourceFileName, 
                                    ref List<TestLog> a_lsTestLogs)
        {
            string sFuncName = "[ReadTestCasesFromTDSFile_Java]";

            // for Excel
            Excel.Worksheet excelSheet;
            Excel.Range excelRange;

            string sClassName = a_sSourceFileName.Replace(".java", "."); // Class name = source file name
            string sMethodName0 = "";
            string sMethodName = "";
            string sTCLabelName0 = "";
            string sTCLabelName = "";
            string sTCFuncName0 = "";
            string sTCFuncName = "";
            string sTCSourceFileName = "";
            string sTCNote = "";
            int dErrorCount = 0;

            TestType eTestType;
            TestMeans eTestMeans;
            TestLog eTestLog = null;

            string sMsgHeader, sMsg;

            Thread searchThread = null;



            string sChapterInSUTS = Constants.StringTokens.ERROR;

            try
            {
                // Get the used range of the "LookupTable" sheet.
                try
                {
                    excelSheet = (Excel.Worksheet)a_excelBook.Worksheets.get_Item(Constants.SHEET_NAME);
                    excelRange = excelSheet.UsedRange;
                }
                catch
                {
                    Logger.Print(Constants.StringTokens.MSG_BULLET, "No \"" + Constants.SHEET_NAME + "\" sheet can be found.");
                    return ++dErrorCount;
                }

                // Check the column count.
                if (4 > excelRange.Columns.Count)
                {
                    Logger.Print(Constants.StringTokens.MSG_BULLET, "Invalid \"" + Constants.SHEET_NAME + "\" sheet.");
                    return ++dErrorCount;
                }


                // Extract the (TC label, TC func) pairs and note.
                int dFirstRow = 2;
                for (int i = dFirstRow; i <= excelRange.Rows.Count; i++)
                {

                    

                    // Ignore the empty row.
                    if ((null == (excelRange.Cells[i, 1] as Excel.Range).Value2) &&
                        (null == (excelRange.Cells[i, 2] as Excel.Range).Value2) &&
                        (null == (excelRange.Cells[i, 3] as Excel.Range).Value2))
                    {
                        if (i == dFirstRow)
                        {
                            Logger.Print(Constants.StringTokens.MSG_BULLET, "No data contained in \"" + Constants.SHEET_NAME + "\" sheet.");
                            dErrorCount++;
                        }
                        break;
                    }

                    
                    sMsgHeader = sFuncName + ":" + Constants.StringTokens.MSG_BULLET + " Row " + i.ToString() + ":";

                    // Read data from the table.
                    int dCol = 1;
                    sMethodName0 = ReadStringFromExcelCell(excelRange.Cells[i, dCol], Constants.StringTokens.DEFAULT_INVALID_VALUE, true);
                    sTCLabelName0 = ReadStringFromExcelCell(excelRange.Cells[i, ++dCol], Constants.StringTokens.DEFAULT_INVALID_VALUE, true);
                    sTCFuncName0 = ReadStringFromExcelCell(excelRange.Cells[i, ++dCol], Constants.StringTokens.DEFAULT_INVALID_VALUE, true);
                    sTCNote = ReadStringFromExcelCell(excelRange.Cells[i, ++dCol], Constants.StringTokens.DEFAULT_INVALID_VALUE, true);

                    // Determine the test means.
                    eTestType = DetermineTestType(sTCNote);
                    eTestMeans = DetermineTestMeans(eTestType);


                    //
                    // Check & modify Note
                    //
                    //if (eTestMeans == TestMeans.TEST_SCRIPT)
                    //{

                    //    if (Constants.StringTokens.NA == sTCFuncName0)
                    //    {

                    //    }
                    //    else
                    //    {
                    //        string c = a_sSourceFileName.Replace(".java", "Test");
                    //        string f = sTCFuncName0 + ".txt";

                    //        var logItems = new Tuple<string, string, List<TestLog>>(c, f, a_lsTestLogs);
                    //        searchThread = new Thread(new ParameterizedThreadStart(SearchTestLog));
                    //        searchThread.Priority = ThreadPriority.AboveNormal;
                    //        //searchThread.IsBackground = true;
                    //        searchThread.Start(logItems);

                    //        //task = SearchTestLogEx;
                    //        //asyncResult = task.BeginInvoke(logItems, null, null);

                    //        //Logger.Print($"Search log {f}", Logger.PrintOption.File);
                    //    }
                    //}



                    sChapterInSUTS = Constants.StringTokens.ERROR;

                    if (a_wordDoc != null && !a_sSourceFileName.StartsWith(Constants.StringTokens.NA))
                    {
                        string className = a_sSourceFileName.Replace(".java", "");
                        string TDSExcelName = a_sTDSFile;

                        var items = new Tuple<Word.Document, string, string>(a_wordDoc, className, TDSExcelName);
                        searchThread = new Thread(new ParameterizedThreadStart(SUTS_FindSectionOfClass_JavaByThread));
                        searchThread.Priority = ThreadPriority.AboveNormal;
                        //searchThread.IsBackground = true;
                        searchThread.Start(items);

                    }


                    // --------------------------------------------------
                    // Check & adjust the read data.
                    // --------------------------------------------------

                    //
                    // Check & modfy method name:
                    //
                    if (sMethodName0.StartsWith(Constants.StringTokens.ERROR_MSG_HEADER))
                    {
                        sMethodName = sMethodName0;
                        Logger.Print(sMsgHeader, ErrorMessage.INVLAID_METHOD_NAME + ": \"" + sMethodName0 + "\"");
                        dErrorCount++;
                    }
                    else if (sMethodName0.StartsWith(Constants.StringTokens.NA))
                    {
                        sMethodName = Constants.StringTokens.ERROR_MSG_HEADER + ErrorMessage.METHOD_NAME_SHALL_NOT_BE_NA;
                        Logger.Print(sMsgHeader, ErrorMessage.METHOD_NAME_SHALL_NOT_BE_NA);
                        dErrorCount++;
                    }
                    else if ("" == sMethodName0)
                    {
                        sMethodName = Constants.StringTokens.ERROR_MSG_HEADER + ErrorMessage.METHOD_NAME_SHALL_NOT_BE_EMPTY;
                        Logger.Print(sMsgHeader, ErrorMessage.METHOD_NAME_SHALL_NOT_BE_EMPTY);
                        dErrorCount++;
                    }
                    else
                    {
                        // Strip the redundant spaces between parameters.
                        sMethodName0 = sMethodName0.Replace(", ", ",");
                        sMethodName0 = sMethodName0.Replace(" ,", ",");

                        if (sMethodName0.Contains(" "))
                        {
                            sMsg = ErrorMessage.METHOD_NAME_SHALL_NOT_CONTAIN_SPACE + ": \"" + sMethodName0 + "\"";
                            sMethodName = Constants.StringTokens.ERROR_MSG_HEADER + sMsg;
                            Logger.Print(sMsgHeader, sMsg);
                            dErrorCount++;
                        }
                        else // Form the unique method name: Filename + method name.
                        {
                            sMethodName = sClassName + ArrangeAndCheckMethodName(sMethodName0);
                        }
                    }


                    //
                    // Check & modify TC label.
                    //
                    if (sTCLabelName0.StartsWith(Constants.StringTokens.ERROR_MSG_HEADER))
                    {
                        sTCLabelName = sTCLabelName0;
                        Logger.Print(sMsgHeader, ErrorMessage.INVLAID_TC_LABEL + ": \"" + sTCLabelName0 + "\"");
                        dErrorCount++;
                    }
                    else if (sTCLabelName0.StartsWith(Constants.StringTokens.NA))
                    {
                        sTCLabelName = Constants.StringTokens.ERROR_MSG_HEADER + ErrorMessage.TC_LABEL_SHALL_NOT_BE_NA;
                        Logger.Print(sMsgHeader, ErrorMessage.TC_LABEL_SHALL_NOT_BE_NA);
                        dErrorCount++;
                    }
                    else if ("" == sTCLabelName0)
                    {
                        sTCLabelName = Constants.StringTokens.ERROR_MSG_HEADER + ErrorMessage.TC_LABEL_SHALL_NOT_BE_EMPTY;
                        Logger.Print(sMsgHeader, ErrorMessage.TC_LABEL_SHALL_NOT_BE_EMPTY);
                        dErrorCount++;
                    }
                    else if (sTCLabelName0.Contains(" "))
                    {
                        sMsg = ErrorMessage.TC_LABEL_SHALL_NOT_CONTAIN_SPACE + ": \"" + sTCLabelName0 + "\"";
                        sTCLabelName = Constants.StringTokens.ERROR_MSG_HEADER + sMsg;
                        Logger.Print(sMsgHeader, sMsg);
                        dErrorCount++;
                    }
                    else
                    {   
                        // Form the unique method name: Filename + TC label name.
                        sTCLabelName = sClassName + sTCLabelName0;
                    }



                    //
                    // Check & modify TC func.
                    //
                    if (sTCFuncName0.StartsWith(Constants.StringTokens.ERROR_MSG_HEADER))
                    {
                        sTCFuncName = sTCFuncName0;
                        Logger.Print(sMsgHeader, ErrorMessage.INVLAID_TC_FUNC_NAME + ": \"" + sTCFuncName0 + "\"");
                        dErrorCount++;
                    }
                    else if ("" == sTCFuncName0)
                    {
                        if (("" == sTCNote) ||
                            (Constants.StringTokens.NA == sTCNote) ||
                            sTCNote.StartsWith(Constants.StringTokens.ERROR_MSG_HEADER))
                        {
                            sTCFuncName = Constants.StringTokens.ERROR_MSG_HEADER + ErrorMessage.NO_TC_FUNC_NAME_CAN_BE_READ;
                            Logger.Print(sMsgHeader, ErrorMessage.NO_TC_FUNC_NAME_CAN_BE_READ);
                            dErrorCount++;
                        }
                        else
                        {
                            sTCFuncName = Constants.StringTokens.NA + " (" + sTCNote + ")";
                        }
                    }
                    else if (Constants.StringTokens.NA == sTCFuncName0)
                    {
                        if (("" == sTCNote) ||
                            (Constants.StringTokens.NA == sTCNote) ||
                            sTCNote.StartsWith(Constants.StringTokens.ERROR_MSG_HEADER))
                        {
                            sTCFuncName = Constants.StringTokens.ERROR_MSG_HEADER + ErrorMessage.REASON_SHALL_BE_GIVEN_FOR_NA_TC_FUNC;
                            Logger.Print(sMsgHeader, ErrorMessage.REASON_SHALL_BE_GIVEN_FOR_NA_TC_FUNC);
                            dErrorCount++;
                        }
                        else if (eTestMeans == TestMeans.TEST_SCRIPT)
                        {
                            sMsg = ErrorMessage.AMBIGUOUS_BETWEEN_TCFUN_TCNOT + ": \"" + sTCFuncName0 + " and " + sTCNote + "\"";
                            sTCFuncName = Constants.StringTokens.ERROR_MSG_HEADER + sMsg;
                            Logger.Print(sMsgHeader, sMsg);
                            dErrorCount++;
                        }
                        else
                        {
                            sTCFuncName = Constants.StringTokens.NA + " (" + sTCNote + ")";
                        }

                    }
                    else if (sTCFuncName0.StartsWith(Constants.StringTokens.NA))
                    {
                        sTCFuncName = sTCFuncName0;

                    }
                    else if (sTCFuncName0.Contains(" "))
                    {
                        sMsg = ErrorMessage.TC_FUNC_NAME_SHALL_NOT_CONTAIN_SPACE + ": \"" + sTCFuncName0 + "\"";
                        sTCFuncName = Constants.StringTokens.ERROR_MSG_HEADER + sMsg;
                        Logger.Print(sMsgHeader, sMsg);
                        dErrorCount++;
                    }
                    else
                    {
                        // Form the unique method name: Filename + method name.
                        sTCFuncName = sClassName + sTCFuncName0;
                    }


  
                    // call eTestLog
                    eTestLog = new TestLog();

                    //
                    // Check & modify Note
                    //
                    if (eTestMeans == TestMeans.TEST_SCRIPT)
                    {

                        if (Constants.StringTokens.NA == sTCFuncName0)
                        {
                            sMsg = ErrorMessage.AMBIGUOUS_BETWEEN_TCFUN_TCNOT + ": \"" + sTCFuncName0 + " and " + sTCNote + "\"";
                            sTCNote = Constants.StringTokens.ERROR_MSG_HEADER + sMsg;
                            Logger.Print(sMsgHeader, sMsg);
                            dErrorCount++;
                        }
                        else
                        {

                            //
                            // find test log
                            //
                            string className = a_sSourceFileName.Replace(".java", "Test");
                            string fileName = sTCFuncName0 + ".txt";


                            Predicate<TestLog> FindValue = delegate (TestLog obj)
                            {
                                return (obj.ClassName == className) && (obj.FileName == fileName);
                            };


                            int index = a_lsTestLogs.FindIndex(FindValue);

                            if (index >= 0)
                            {
                                a_lsTestLogs[index].Increment();
                                eTestLog = a_lsTestLogs[index];
                            }



                            //while (!asyncResult.IsCompleted)
                            //{
                            //    Thread.SpinWait(0);
                            //}
                            //// 
                            //int index = task.EndInvoke(asyncResult);

                            //searchThread.Join();

                            //int index = gLogIndex;
                            //if (index >= 0)
                            //{
                            //    a_lsTestLogs[index].Increment();
                            //    eTestLog = a_lsTestLogs[index];

                            //    //Logger.Print($"Found log {eTestLog.FileName}", Logger.PrintOption.File);
                            //}
                        }
                    }

                    //
                    // Find SUTS
                    //
                    //sChapterInSUTS = Constants.StringTokens.ERROR;

                    //if (a_wordDoc != null && !a_sSourceFileName.StartsWith(Constants.StringTokens.NA))
                    //{
                    //    string className = a_sSourceFileName.Replace(".java", "");
                    //    string docName = a_sTDSFile;
                    //    sChapterInSUTS = SUTS_FindSectionOfClass_Java(a_wordDoc, className, docName);
                    //}


                    // Determine the test case source file name.
                    if ((sTCFuncName0.StartsWith(Constants.StringTokens.NA) == true) &&
                        (sTCNote.StartsWith(Constants.StringTokens.ERROR_MSG_HEADER) == false))
                    {
                        sTCSourceFileName = sTCFuncName;
                    }
                    else
                    {
                        sTCSourceFileName = a_sSourceFileName.Replace(".java", "Test.java");
                    }


                    searchThread.Join();
                    sChapterInSUTS = gSUTSChapter;


                    //
                    // Record the data
                    //
                    TestCaseItem tItem = new TestCaseItem();

                    tItem.sSourceFileName = a_sSourceFileName;
                    tItem.sMethodName = sMethodName;
                    tItem.sTCLabelName = sTCLabelName;
                    tItem.sTCFuncName = sTCFuncName;
                    tItem.sTDSFileName = a_sTDSFile;
                    tItem.sTCSourceFileName = sTCSourceFileName;
                    tItem.sTCNote = sTCNote;
                    tItem.eTestMeans = eTestMeans;
                    tItem.eTestlog = eTestLog;
                    tItem.eType = eTestType;
                    tItem.sChapterInSUTS = sChapterInSUTS;
                    
                    g_tTestCaseTable.ltItems.Add(tItem);

                }

            }
            catch (SystemException ex)
            {
                Logger.Print(sFuncName, ex.ToString());
                dErrorCount++;
            }

            return dErrorCount;
        }


        /// <summary>
        /// Determin the test type by TCNote
        /// </summary>
        /// <param name="a_sInfo"></param>
        /// <returns></returns>
        private TestType DetermineTestType(string a_sNote)
        {
            TestType eTestType = TestType.Unknow;

            if (GetStringValue(TestType.ByMockito).Equals(a_sNote))
            {
                eTestType = TestType.ByMockito;
            }
            else if (GetStringValue(TestType.ByPowerMockito).Equals(a_sNote))
            {
                eTestType = TestType.ByPowerMockito;
            }


            else if (GetStringValue(TestType.GetterSetter).Equals(a_sNote))
            {
                eTestType = TestType.GetterSetter;
            }
            else if (GetStringValue(TestType.Empty).Equals(a_sNote))
            {
                eTestType = TestType.Empty;
            }
            else if (GetStringValue(TestType.Abstract).Equals(a_sNote))
            {
                eTestType = TestType.Abstract;
            }
            else if (GetStringValue(TestType.Interface).Equals(a_sNote))
            {
                eTestType = TestType.Interface;
            }
            else if (GetStringValue(TestType.Native).Equals(a_sNote))
            {
                eTestType = TestType.Native;
            }

            else if (GetStringValue(TestType.ByCodeAnalysis).Equals(a_sNote))
            {
                eTestType = TestType.ByCodeAnalysis;
            }
            else if (GetStringValue(TestType.PureFunctionCalls).Equals(a_sNote))
            {
                eTestType = TestType.PureFunctionCalls;
            }
            else if (GetStringValue(TestType.PureUIFunctionCalls).Equals(a_sNote))
            {
                eTestType = TestType.PureUIFunctionCalls;
            }
            else
            {
                eTestType = TestType.Unknow;
            }

            return eTestType;
        }


        /// <summary>
        /// Determine the meaning of test type.
        /// </summary>
        /// <param name="testType"></param>
        /// <returns></returns>
        private TestMeans DetermineTestMeans(TestType testType)
        {
            TestMeans eTestMeans = TestMeans.UNKNOWN;

            switch (testType)
            {
                case TestType.ByMockito:
                case TestType.ByPowerMockito:
                    eTestMeans = TestMeans.TEST_SCRIPT;
                    break;

                case TestType.GetterSetter:
                case TestType.Empty:
                case TestType.Abstract:
                case TestType.Interface:
                case TestType.Native:
                    eTestMeans = TestMeans.NA;
                    break;


                case TestType.PureFunctionCalls:
                case TestType.PureUIFunctionCalls:
                case TestType.ByCodeAnalysis:
                    eTestMeans = TestMeans.CODE_ANALYSIS;
                    break;

                default:
                    eTestMeans = TestMeans.UNKNOWN;
                    break;
            }

            return eTestMeans;
        }


        #endregion













        #region C Group



        /// <summary>
        /// 
        /// </summary>
        /// <param name="a_excelBook"></param>
        /// <param name="a_wordDoc"></param>
        /// <param name="a_sTDSFile"></param>
        /// <param name="a_sSourceFileName"></param>
        /// <param name="a_sMethodName"></param>
        /// <param name="a_lsTestLogs"></param>
        /// <returns></returns>
        private int ReadTestCasesFromTDSFile_C(
                                Excel.Workbook a_excelBook,
                                Word.Document a_wordDoc, 
                                ref string a_sTDSFile, 
                                ref string a_sSourceFileName, 
                                ref string a_sMethodName, 
                                ref List<TestLog> a_lsTestLogs)
        {
            const string sFuncName = "[ReadTestCasesFromTDSFile_C]";

            Excel.Worksheet excelSheet;
            Excel.Range excelRange;

            int dRow;
            int dCol;
            string sMethodName = "";
            string sTCLabelName = "";
            int dErrorCount = 0;
            string sMsgHeader, sMsg;


            // -------------------------------------------------------------------------
            // Read data form the "TestCase" sheet.
            // -------------------------------------------------------------------------
            try
            {
                // Get the used range of the "TestCase" sheet.
                try
                {
                    excelSheet = (Excel.Worksheet)a_excelBook.Worksheets.get_Item("TestCase");
                    excelRange = excelSheet.UsedRange;
                }
                catch
                {
                    Logger.Print(Constants.StringTokens.MSG_BULLET, "No \"TestCase\" sheet can be found.");
                    return ++dErrorCount;
                }


                // Calibrate the (0, 0) location.
                bool bFound = false;
                int dRowOffset = 0;
                int dColOffset = 0;
                for (int i = 1; i <= 2; i++)
                {
                    for (int j = 1; j <= 2; j++)
                    {
                        string sValue = ReadStringFromExcelCell(excelRange.Cells[i, j], "", true);
                        if (sValue == "UUT ¢")
                        {
                            dRowOffset = i - 1;
                            dColOffset = j - 1;

                            bFound = true;
                            break;
                        }
                    }
                    if (bFound)
                    {
                        break;
                    }
                }


                // Extract the method name.
                dRow = dRowOffset + 1;
                dCol = dColOffset + 2;
                sMethodName = ReadStringFromExcelCell(excelRange.Cells[dRow, dCol], Constants.StringTokens.DEFAULT_INVALID_VALUE, true);

                sMsgHeader = Constants.StringTokens.MSG_BULLET + " Cell(" + dRow.ToString() + "," + dCol.ToString() + "):";

                // Check the extracted name.
                if (sMethodName.StartsWith(Constants.StringTokens.ERROR_MSG_HEADER))
                {
                    Logger.Print(sMsgHeader, ErrorMessage.INVLAID_METHOD_NAME + ": \"" + sMethodName + "\"");
                    return ++dErrorCount;
                }
                else if (sMethodName.StartsWith(Constants.StringTokens.NA))
                {
                    sMethodName = Constants.StringTokens.ERROR_MSG_HEADER + ErrorMessage.METHOD_NAME_SHALL_NOT_BE_NA;
                    Logger.Print(sMsgHeader, ErrorMessage.METHOD_NAME_SHALL_NOT_BE_NA);
                    return ++dErrorCount;
                }
                else if ("" == sMethodName)
                {
                    sMethodName = Constants.StringTokens.ERROR_MSG_HEADER + ErrorMessage.METHOD_NAME_SHALL_NOT_BE_EMPTY;
                    Logger.Print(sMsgHeader, ErrorMessage.METHOD_NAME_SHALL_NOT_BE_EMPTY);
                    return ++dErrorCount;
                }
                else if (sMethodName.Contains(" "))
                {
                    sMsg = ErrorMessage.METHOD_NAME_SHALL_NOT_CONTAIN_SPACE + ": \"" + sMethodName + "\"";
                    sMethodName = Constants.StringTokens.ERROR_MSG_HEADER + sMsg;
                    Logger.Print(sMsgHeader, sMsg);
                    return ++dErrorCount;
                }
                // Arrange the name, if needs.
                else if (sMethodName.Contains("::"))
                    sMethodName = sMethodName.Replace("::", ".");

                // Check the consistence of method names.
                if (!a_sMethodName.Contains(sMethodName))
                {
                    Logger.Print(sMsgHeader, "method name \"" + sMethodName + "\" is not consisted with the name \"" + a_sMethodName + "\" in the cover sheet.");
                    return ++dErrorCount;
                }

                // Extract TC label names.
                dRow = dRowOffset + 3;
                dCol = dColOffset + 2;
                string sPrevLabel = "";
                for (int i = dCol; i <= excelRange.Columns.Count; i++)
                {
                    // Read TC label name.
                    sTCLabelName = ReadStringFromExcelCell(excelRange.Cells[dRow, i], "", true);

                    // Arrange the name, if needs.
                    if (sTCLabelName.Contains("::"))
                        sTCLabelName = sTCLabelName.Replace("::", ".");

                    // skip empty and covered item
                    if (("" == sTCLabelName) || sTCLabelName.StartsWith("Covered by"))
                    {
                        continue;
                    }

                    sMsgHeader = Constants.StringTokens.MSG_BULLET + " Cell(" + dRow.ToString() + "," + i.ToString() + "):";

                    // Check the extracted name.
                    if (sTCLabelName.StartsWith(Constants.StringTokens.ERROR_MSG_HEADER))
                    {
                        Logger.Print(sMsgHeader, "Invlaid TC label read: \"" + sTCLabelName + "\"");
                        dErrorCount++;
                    }
                    else if (sTCLabelName.StartsWith(Constants.StringTokens.NA))
                    {
                        sTCLabelName = Constants.StringTokens.ERROR_MSG_HEADER + ErrorMessage.TC_LABEL_SHALL_NOT_BE_NA;
                        Logger.Print(sMsgHeader, ErrorMessage.TC_LABEL_SHALL_NOT_BE_NA);
                        dErrorCount++;
                    }
                    else if (sTCLabelName.Contains(" "))
                    {
                        sMsg = ErrorMessage.TC_LABEL_SHALL_NOT_CONTAIN_SPACE + ": \"" + sTCLabelName + "\"";
                        sTCLabelName = Constants.StringTokens.ERROR_MSG_HEADER + sMsg;
                        Logger.Print(sMsgHeader, sMsg);
                        dErrorCount++;
                    }
                    // Skip this item if TC label is same as the previous one.
                    else if (sPrevLabel == sTCLabelName)
                    {
                        Logger.Print(sMsgHeader, ErrorMessage.DUPLICATE_TC_LABEL_FOUND + ": \"" + sTCLabelName + "\"");
                        dErrorCount++;
                        continue;
                    }


                    // new a objec TestLog for init.
                    TestLog eTestLog = new TestLog();
                    string fileName = a_sMethodName + "." + sTCLabelName + ".txt";


                    Predicate<TestLog> FindValue = delegate (TestLog obj)
                    {
                        return obj.FileName == fileName;
                    };


                    int index = a_lsTestLogs.FindIndex(FindValue);

                    if (index >= 0)
                    {
                        a_lsTestLogs[index].Increment();
                        eTestLog = a_lsTestLogs[index];
                    }




                    //
                    // Find SUTS
                    //
                    string sChapterInSUTS = "";

                    if (!a_sSourceFileName.StartsWith(Constants.StringTokens.NA))
                    { 
                        
                        sChapterInSUTS = SUTS_FindSectionOfClass_C(a_wordDoc, a_sMethodName, sTCLabelName);
                    }



                    // Record the data read.
                    TestCaseItem tItem = new TestCaseItem();

                    tItem.sSourceFileName = a_sSourceFileName;
                    tItem.sMethodName = a_sMethodName;
                    tItem.sTCSourceFileName = Constants.StringTokens.NA;
                    tItem.sTCLabelName = sTCLabelName.StartsWith(Constants.StringTokens.ERROR_MSG_HEADER) ?
                        sTCLabelName : a_sMethodName + "." + sTCLabelName;
                    tItem.sTCFuncName = Constants.StringTokens.NA;
                    tItem.sTDSFileName = a_sTDSFile;
                    tItem.sTCNote = Constants.StringTokens.NA;
                    tItem.eTestMeans = TestMeans.TEST_SCRIPT;
                    tItem.eTestlog = eTestLog;
                    tItem.eType = TestType.ByVectorCast;
                    tItem.sChapterInSUTS = sChapterInSUTS;

                    g_tTestCaseTable.ltItems.Add(tItem);

                    sPrevLabel = sTCLabelName;
                }
            }
            catch (SystemException ex)
            {
                Logger.Print(sFuncName, ex.ToString());
                dErrorCount++;
            }

            return dErrorCount;
        }

        #endregion


    }
}
