using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;


namespace UTChecker
{
    public partial class UTChecker
    {


        /// <summary>
        /// 
        /// </summary>
        private void ReleaseOfficeApps()
        {
            string sFuncName = "[ReleaseOfficeApps]";

            try
            {
                // Close EXCEL app;
                if (null != g_excelApp)
                {
                    g_excelApp.DisplayAlerts = false;
                    g_excelApp.Quit();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(g_excelApp); // force the app to be closed
                    g_excelApp = null;
                }
            }
            catch (SystemException e)
            {
                Logger.Print(sFuncName, "Cannot release Excel app. Please kill the app manually: " + e.ToString());
            }
        }



        /// <summary>
        /// 
        /// </summary>
        /// <param name="a_excelApp"></param>
        /// <param name="a_sExcelFile"></param>
        /// <param name="a_bReadOnly"></param>
        /// <returns></returns>
        private Excel.Workbook OpenExcelWorkbook(Excel.Application a_excelApp, string a_sExcelFile, bool a_bReadOnly)
        {
            string sFuncName = "[OpenExcelWorkbook]";

            if (null == a_excelApp)
            {
                Logger.Print(sFuncName, "Null EXCEL app is given.");
                return null;
            }

            try
            {
                Excel.Workbook excelBook = a_excelApp.Workbooks.Open(a_sExcelFile, 0, a_bReadOnly, 6, "", "",
                    true, Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                return excelBook;
            }
            catch (System.Exception ex)
            {
                Logger.Print(sFuncName, "Open EXCEL file failed: " + a_sExcelFile + " (" + ex.ToString() + ")");
                return null;
            }
        }



        /// <summary>
        /// 
        /// </summary>
        /// <param name="a_excelBook"></param>
        /// <param name="a_sSheetName"></param>
        /// <returns></returns>
        private Excel.Worksheet GetExcelSheet(Excel.Workbook a_excelBook, string a_sSheetName)
        {
            string sFuncName = "[GetExcelSheet]";

            if (null == a_excelBook)
            {
                Logger.Print(sFuncName, "Null EXCEL workbook is given.");
                return null;
            }

            try
            {
                Excel.Worksheet excelSheet = (Excel.Worksheet)a_excelBook.Worksheets.get_Item(a_sSheetName);
                return excelSheet;
            }
            catch (System.Exception ex)
            {
                Logger.Print(sFuncName, "Cannot find EXCEL sheet: " + a_sSheetName + " (" + ex.ToString() + ")");
                return null;
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="a_excelCell"></param>
        /// <param name="a_sNullStringToBeUsed"></param>
        /// <param name="a_bStripSpaces"></param>
        /// <returns></returns>
        private string ReadStringFromExcelCell(Excel.Range a_excelCell, string a_sNullStringToBeUsed, bool a_bStripSpaces)
        {
            string sValue;

            //excelCell = excelRange.Cells[i, SWDDLookupTable.COLUMN_DESIGN_ID] as Excel.Range;
            if (null == a_excelCell.Value2)
                sValue = a_sNullStringToBeUsed;
            else
            {
                // Try to get the data as a string. 
                try
                {
                    sValue = (string)a_excelCell.Value2;
                }
                // Otherwise, try to get the data as a double.
                catch
                {
                    try
                    {
                        double fValue = a_excelCell.Value2;
                        sValue = fValue.ToString();
                    }
                    catch
                    {
                        sValue = a_sNullStringToBeUsed;
                    }
                }

                // Strip leading/tailing spaces, if needs.
                if (a_bStripSpaces)
                    sValue = sValue.Trim();
            }

            return sValue;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="a_excelCell"></param>
        /// <param name="a_dNullValueToBeUsed"></param>
        /// <returns></returns>
        private int ReadValueFromExcelCell(Excel.Range a_excelCell, int a_dNullValueToBeUsed)
        {
            int dValue;

            if (null == a_excelCell.Value2)
                dValue = a_dNullValueToBeUsed;
            else
            {
                // Try to get the data as an integer. 
                try
                {
                    dValue = (int)a_excelCell.Value2;
                }
                // Otherwise, try to get the data as a double.
                catch
                {
                    try
                    {
                        double fValue = a_excelCell.Value2;
                        dValue = (int)(fValue + 0.5);
                    }
                    // Otherwise, try to get the data as a string.
                    catch
                    {
                        try
                        {
                            string sValue = (string)a_excelCell.Value2;
                            dValue = Convert.ToInt32(sValue);
                        }
                        catch
                        {
                            dValue = a_dNullValueToBeUsed;
                        }
                    }
                }
            }

            return dValue;
        }


        /// <summary>
        /// ExtractMethodName
        /// </summary>
        /// <param name="a_sLine"></param>
        /// <returns></returns>
        private string ExtractMethodName(string a_sLine)
        {
            if ("" == a_sLine)
                return "";

            if (a_sLine.Contains("::"))
                a_sLine = a_sLine.Replace("::", ".");

            // Strip the "(...) part.
            // Before: "XXX(...)"
            // After:  "XXX"
            if (a_sLine.Contains("("))
            {
                int dPosition = a_sLine.IndexOf('(');
                a_sLine = a_sLine.Substring(0, dPosition);
            }

            // Return the last part.
            // Before: "public void XXX"
            // After:  "XXX"
            if (a_sLine.Contains(" "))
            {
                string[] asElements = a_sLine.Split(' ');

                for (int i = asElements.Length - 1; i >= 0; i--)
                {
                    if ("" != asElements[i])
                    {
                        if (asElements[i].StartsWith("*"))
                            return asElements[i].Substring(1, asElements[i].Length - 1);
                        else
                            return asElements[i];
                    }
                }

                return "";
            }
            else
                return a_sLine;
        }


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
                    break;
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
        /// <returns></returns>
        private bool ReadDataFromTDSFiles(string a_sModuleName, ref List<string> a_lsTDSFiles, ref List<TestLog> a_lsTestLogs)
        {
            string sFuncName = "[ReadDataFromTDSFiles]";

            Excel.Workbook excelBook = null;
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
                Logger.Print(sFuncName, "EXCEL app is null.");
                return false;
            }

            try
            {
                Logger.Print(sFuncName, "Reading TDS files...");

                // Initialize objects.
                dErrorCount = 0;
                g_tTestCaseTable.ltItems.Clear();
                g_excelApp.DisplayAlerts = false; // show no alert while closing the file

                // Read data from each TDS file.
                foreach (string sFile in a_lsTDSFiles)
                {
                    sFileNameWithoutPath = "\"" + Path.GetFileName(sFile) + "\"";
                    Logger.Print("", "Reading " + sFileNameWithoutPath + "...");

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
                            ReadTestCasesFromTDSFile_Java(excelBook, ref sShortTDSFileName, ref sSourceFileName, ref a_lsTestLogs);
                        }
                        else
                        {
                            ReadTestCasesFromTDSFile_C(excelBook, ref sShortTDSFileName, ref sSourceFileName, ref sMethodName);
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
                        // Close the TDS file.
                        excelBook.Close();
                    }
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


        private int CompareTDSandTestLogs(string name, string templateFile, List<string> testLogs)
        {
            string sFuncName = "[CompareTDSandTestLogs]";
            string compLog = name + "_cmp.txt";

            int errorCount = -1;

            // for excel workbook
            Excel.Workbook excelBook;

            // for sheet LookupTable
            Excel.Worksheet excelSheetLookupTable;
            Excel.Range excelRangeLookupTable;

            // for sheet Summary
            Excel.Worksheet excelSheetSummary;
            Excel.Range excelRangeSummary;



            List<string> tempLog = testLogs;

            // Check the EXCEL app.
            if (null == g_excelApp)
            {
                Logger.Print(sFuncName, "EXCEL app is null.");
                return -1;
            }


            // Open the TDS file & get the lookup-table sheet.
            //g_excelApp.DisplayAlerts = false;

            excelBook = OpenExcelWorkbook(g_excelApp, templateFile, false);  // ready only
            if (null == excelBook)
            {
                Logger.Print($"{sFuncName} Error occurred when open excel file.");
                return -1;
            }



            // Get the used range of the "LookupTable" sheet.
            try
            {
                excelSheetLookupTable = (Excel.Worksheet)excelBook.Worksheets.get_Item(Constants.SHEET_NAME);
                excelRangeLookupTable = excelSheetLookupTable.UsedRange;
                

                excelSheetSummary = (Excel.Worksheet)excelBook.Worksheets.get_Item(Constants.SHEET_SUMMARY);
                excelRangeSummary = excelSheetSummary.UsedRange;

            }
            catch
            {
                Logger.Print(Constants.StringTokens.MSG_BULLET, "No \"" + Constants.SHEET_NAME + "\" sheet can be found.");
                return -1;
            }



            // clear
            List<string> testItem = new List<string>();
            testItem.Clear();

            // clear
            List<string> noNeedItem = new List<string>();
            noNeedItem.Clear();



            int dFirstRow = 3;
            for (int i = dFirstRow; i <= excelRangeLookupTable.Rows.Count; i++)
            {

                string NG_Mark= ReadStringFromExcelCell(excelRangeLookupTable.Cells[i, TestCaseTableConstants.ColumnIndex.NG_MARKER], "", true);

                if (NG_Mark.Equals("X"))
                {
                    continue;
                }


                // get the class name from excel column B
                string className = ReadStringFromExcelCell(excelRangeLookupTable.Cells[i, TestCaseTableConstants.ColumnIndex.SOURCE_FILE], Constants.StringTokens.DEFAULT_INVALID_VALUE, true);

                if (className.IndexOf(".java") == -1)
                {
                    continue;
                }

                // remove ext filename ".java"
                className = className.Replace(".java", "");

                string testCaseSouce = ReadStringFromExcelCell(excelRangeLookupTable.Cells[i, TestCaseTableConstants.ColumnIndex.TC_SOURCE_FILE], Constants.StringTokens.DEFAULT_INVALID_VALUE, true);
                string testCaseName = ReadStringFromExcelCell(excelRangeLookupTable.Cells[i, TestCaseTableConstants.ColumnIndex.TC_NAME], Constants.StringTokens.DEFAULT_INVALID_VALUE, true);
                string testType = ReadStringFromExcelCell(excelRangeLookupTable.Cells[i, TestCaseTableConstants.ColumnIndex.NOTE], Constants.StringTokens.DEFAULT_INVALID_VALUE, true);


                if (testType.Equals(TestType.ByPowerMocktio) || testType.Equals(TestType.ByMockito))
                {

                    // get the test case source from excel column G
                    testCaseSouce = testCaseSouce.Replace(".java", "");
                    testCaseName = testCaseName.Substring(className.Length + 1); // with '.', so plus 1


                    
                    StringBuilder tc_className = new StringBuilder("");
                    StringBuilder tc_fileName = new StringBuilder("");
                    StringBuilder tc_path = new StringBuilder("");


                    int testLogsIndex = 0;
                    bool bFind = false;

                    foreach (string line in tempLog)
                    {
                        char[] spilter = new char[] { ',', ' ' };
                        string[] s = line.Split(spilter, StringSplitOptions.RemoveEmptyEntries);

                        tc_className.Clear();
                        tc_fileName.Clear();
                        tc_path.Clear();

                        tc_className.Append(s[0]);
                        tc_fileName.Append(s[1]);
                        tc_path.Append(s[2]);

                        if (tc_className.ToString().Equals(testCaseSouce) &&
                            tc_fileName.ToString().Equals(testCaseName + ".txt"))
                        {
                            bFind = true;
                            break;
                        }

                        testLogsIndex++;

                    }

                    if (bFind)
                    {
                        // remove founded item from list.
                        tempLog.RemoveAt(testLogsIndex);

                        excelRangeLookupTable.Cells[i, TestCaseTableConstants.ColumnIndex.TEST_LOG] = tc_fileName.ToString();
                        excelRangeLookupTable.Cells[i, TestCaseTableConstants.ColumnIndex.TEST_LOG_PATH] = tc_path.ToString();

                    }

                    testItem.Add($"{i}, {className}, {testCaseSouce}, {testCaseName}, {tc_fileName.ToString()}");

                }
                else
                {
                    excelRangeLookupTable.Cells[i, TestCaseTableConstants.ColumnIndex.TEST_LOG] = "N/A";
                    excelRangeLookupTable.Cells[i, TestCaseTableConstants.ColumnIndex.TEST_LOG_PATH] = "N/A";

                    noNeedItem.Add($"{i}, {className}, {testCaseSouce}, {testCaseName}");
                }

            }


            //if (File.Exists(compLog))
            //{
            //    File.Delete(compLog);
            //}

            //foreach (string s in testItem)
            //{
            //    using (StreamWriter sw = File.AppendText(compLog))
            //    {
            //        sw.WriteLine(s);
            //    }
            //}

            //foreach (string s in noNeedItem)
            //{
            //    using (StreamWriter sw = File.AppendText(compLog))
            //    {
            //        sw.WriteLine(s);
            //    }
            //}

            if (tempLog.Count != 0)
            {
                int lastLine = excelRangeLookupTable.Rows.Count + 1;

                foreach (string line in tempLog)
                {
                    //using (StreamWriter sw = File.AppendText(compLog))
                    //{
                    //    sw.WriteLine($"Unknow: {line}");
                    //}

                    char[] spilter = new char[] { ',', ' ' };
                    string[] s = line.Split(spilter, StringSplitOptions.RemoveEmptyEntries);

                    string testCodeClass = s[0];
                    string testCodeFileName = s[1];
                    string testCaseName = s[1].Replace(".txt", "");
                    string testCodePath = s[2];

                    excelRangeLookupTable.Cells[lastLine, TestCaseTableConstants.ColumnIndex.NG_MARKER] = Constants.StringTokens.X;
                    excelRangeLookupTable.Cells[lastLine, TestCaseTableConstants.ColumnIndex.SOURCE_FILE] = "N/A";
                    excelRangeLookupTable.Cells[lastLine, TestCaseTableConstants.ColumnIndex.METHOD_NAME] = "N/A";
                    excelRangeLookupTable.Cells[lastLine, TestCaseTableConstants.ColumnIndex.TC_LABEL] = "N/A";
                    excelRangeLookupTable.Cells[lastLine, TestCaseTableConstants.ColumnIndex.TC_NAME] = "N/A";
                    excelRangeLookupTable.Cells[lastLine, TestCaseTableConstants.ColumnIndex.TDS_FILE] = "N/A";
                    excelRangeLookupTable.Cells[lastLine, TestCaseTableConstants.ColumnIndex.TC_SOURCE_FILE] = testCodeClass + ".java";
                    excelRangeLookupTable.Cells[lastLine, TestCaseTableConstants.ColumnIndex.NOTE] = "N/A";
                    excelRangeLookupTable.Cells[lastLine, TestCaseTableConstants.ColumnIndex.TEST_LOG] = testCaseName;
                    excelRangeLookupTable.Cells[lastLine, TestCaseTableConstants.ColumnIndex.TEST_LOG_PATH] = testCodePath;
                    lastLine++;
                }

                // update the error count in sheet "Summary"
                string errString = ReadStringFromExcelCell(excelRangeSummary.Cells[TestCaseTableConstants.RowIndex.ERROR_COUNT, 2], "", true);

                excelRangeSummary.Cells[TestCaseTableConstants.RowIndex.ERROR_COUNT, 2] = Convert.ToInt16(errString) + tempLog.Count; ;

                errorCount = tempLog.Count;


                try
                {
                    // autofilter
                    excelSheetLookupTable.Range["A1", "A1"].AutoFilter(1, "X");

                }
                catch (SystemException e)
                {
                    Logger.Print(e.ToString());
                }

            }
            else
            {
                errorCount = 0;
            }

            excelBook.Save();
            excelBook.Close();


            return errorCount;
        }



    }
}
