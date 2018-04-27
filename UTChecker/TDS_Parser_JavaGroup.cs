using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;


namespace UTChecker
{
    public partial class TDS_Parser
    {


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
                    excelSheet = (Excel.Worksheet)a_excelBook.Worksheets.get_Item(TestCaseTableConstants.SHEET_NAME);
                    excelRange = excelSheet.UsedRange;
                }
                catch
                {
                    Logger.Print(Constants.StringTokens.MSG_BULLET, "No \"" + TestCaseTableConstants.SHEET_NAME + "\" sheet can be found.");
                    return ++dErrorCount;
                }

                // Check the column count.
                if (4 > excelRange.Columns.Count)
                {
                    Logger.Print(Constants.StringTokens.MSG_BULLET, "Invalid \"" + TestCaseTableConstants.SHEET_NAME + "\" sheet.");
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
                            Logger.Print(Constants.StringTokens.MSG_BULLET, "No data contained in \"" + TestCaseTableConstants.SHEET_NAME + "\" sheet.");
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
                    else // Form the unique method name: Filename + TC label name.
                        sTCLabelName = sClassName + sTCLabelName0;


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
                            sTCFuncName = Constants.StringTokens.NA + " (" + sTCNote + ")";
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
                            sTCFuncName = Constants.StringTokens.NA + " (" + sTCNote + ")";
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
                    else // Form the unique method name: Filename + method name.
                        sTCFuncName = sClassName + sTCFuncName0;

                    // Determine the test case source file name.
                    if (sTCFuncName0.StartsWith(Constants.StringTokens.NA))
                        sTCSourceFileName = sTCFuncName;
                    else
                        sTCSourceFileName = a_sSourceFileName.Replace(".java", "Test.java");

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



    }
}
