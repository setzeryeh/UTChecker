using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;


namespace UTChecker
{
    public partial class UTChecker
    {


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
                    return "";
                else
                    return a_sOrgMethodName;
            }
        }


        /// <summary>
        /// 
        /// </summary>
        /// <param name="a_excelBook"></param>
        /// <param name="a_sTDSFile"></param>
        /// <param name="a_sSourceFileName"></param>
        /// <returns></returns>
        private int ReadTestCasesFromTDSFile_Java(Excel.Workbook a_excelBook, ref string a_sTDSFile, ref string a_sSourceFileName)
        {
            string sFuncName = "[ReadTestCasesFromTDSFile_Java]";

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
            TestMeans eTestMeans;

            string sMsgHeader, sMsg;


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

                    sMsgHeader = Constants.StringTokens.MSG_BULLET + " Row " + i.ToString() + ":";

                    // Read data from the table.
                    int dCol = 1;
                    sMethodName0 = ReadStringFromExcelCell(excelRange.Cells[i, dCol], Constants.StringTokens.DEFAULT_INVALID_VALUE, true);
                    sTCLabelName0 = ReadStringFromExcelCell(excelRange.Cells[i, ++dCol], Constants.StringTokens.DEFAULT_INVALID_VALUE, true);
                    sTCFuncName0 = ReadStringFromExcelCell(excelRange.Cells[i, ++dCol], Constants.StringTokens.DEFAULT_INVALID_VALUE, true);
                    sTCNote = ReadStringFromExcelCell(excelRange.Cells[i, ++dCol], Constants.StringTokens.DEFAULT_INVALID_VALUE, true);

                    // Determine the test means.
                    eTestMeans = DetermineTestMeans(sTCNote);

                    // --------------------------------------------------
                    // Check & adjust the read data.
                    // --------------------------------------------------
                    // Check & modfy method name:
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
                            sMethodName = sClassName + ArrangeAndCheckMethodName(sMethodName0);
                    }

                    // Check & modify TC label.
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


                    // Check & modify TC func.
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

                    if (eTestMeans == TestMeans.TEST_SCRIPT && sTCFuncName0.StartsWith(Constants.StringTokens.NA))
                    {
                        sMsg = ErrorMessage.MISSING_TESTCASE_FUNCTION_NAME + ": \"" + sTCFuncName0 + "\"";
                        sTCFuncName = Constants.StringTokens.ERROR_MSG_HEADER + sMsg;
                        Logger.Print(sMsgHeader, sMsg);
                        dErrorCount++;
                    }




                    // Determine the test case source file name.
                    if (sTCFuncName0.StartsWith(Constants.StringTokens.NA))
                    {
                        sTCSourceFileName = sTCFuncName;
                    }
                    else
                    {
                        sTCSourceFileName = a_sSourceFileName.Replace(".java", "Test.java");
                    }





                    // Record the data read.
                    TestCaseItem tItem = new TestCaseItem();
                    tItem.sTDSFileName = a_sTDSFile;
                    tItem.sSourceFileName = a_sSourceFileName;
                    tItem.sMethodName = sMethodName;
                    tItem.sTCLabelName = sTCLabelName;
                    tItem.sTCFuncName = sTCFuncName;
                    tItem.sTCSourceFileName = sTCSourceFileName;
                    tItem.sTCNote = sTCNote;
                    tItem.eTestMeans = eTestMeans;
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

                Logger.Print(" - UNKNOW: ", String.Format("\"{0}\"", a_sInfo));
            }

            return eTestMeans;
        }


        #endregion




        #region C Group



        /// <summary>
        /// 
        /// </summary>
        /// <param name="a_excelBook"></param>
        /// <param name="a_sTDSFile"></param>
        /// <param name="a_sSourceFileName"></param>
        /// <param name="a_sMethodName"></param>
        /// <returns></returns>
        private int ReadTestCasesFromTDSFile_C(Excel.Workbook a_excelBook, ref string a_sTDSFile, ref string a_sSourceFileName, ref string a_sMethodName)
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
                        break;
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
                        continue;

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

                    // Record the data read.
                    TestCaseItem tItem = new TestCaseItem();
                    tItem.sTDSFileName = a_sTDSFile;
                    tItem.sSourceFileName = a_sSourceFileName;
                    tItem.sMethodName = a_sMethodName;
                    tItem.sTCSourceFileName = Constants.StringTokens.NA;
                    tItem.sTCLabelName = sTCLabelName.StartsWith(Constants.StringTokens.ERROR_MSG_HEADER) ?
                        sTCLabelName : a_sMethodName + "." + sTCLabelName;
                    tItem.sTCFuncName = Constants.StringTokens.NA;
                    tItem.sTCNote = Constants.StringTokens.NA;
                    tItem.eTestMeans = TestMeans.TEST_SCRIPT;
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
