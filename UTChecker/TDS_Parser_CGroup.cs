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

    }
}
