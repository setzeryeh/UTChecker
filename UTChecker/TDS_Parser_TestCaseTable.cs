using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace UTChecker
{
    public partial class TDS_Parser
    {

        public bool WriteSummarySheet(Excel.Workbook a_excelBook, string a_sOutFile)
        {
            string sFuncName = "[WriteSummarySheet]";

            Excel.Worksheet excelSheet;
            Excel.Range excelRange;

            try
            {
                // Get the used range of the 1st sheet.
                excelSheet = (Excel.Worksheet)a_excelBook.Worksheets.get_Item(1);
                excelRange = excelSheet.UsedRange;

                int dCol = 2;
                excelRange.Cells[TestCaseTableConstants.RowIndex.DOC_NAME, dCol] = Path.GetFileName(a_sOutFile);
                excelRange.Cells[TestCaseTableConstants.RowIndex.DATE_TIME, dCol] = System.DateTime.Now;

                excelRange.Cells[TestCaseTableConstants.RowIndex.SOURCE_FILE_COUNT, dCol] = g_tTestCaseTable.dSourceFileCount;
                excelRange.Cells[TestCaseTableConstants.RowIndex.METHOD_COUNT, dCol] = g_tTestCaseTable.dMethodCount;
                excelRange.Cells[TestCaseTableConstants.RowIndex.TC_COUNT, dCol] = g_tTestCaseTable.dNormalEntryCount;

                // Write test case test type info.
                excelRange.Cells[TestCaseTableConstants.RowIndex.TC_TEST_VIA_NA_COUNT, dCol] = g_tTestCaseTable.dByNACount;
                excelRange.Cells[TestCaseTableConstants.RowIndex.TC_TEST_VIA_SCRIPT_COUNT, dCol] = g_tTestCaseTable.dByTestScriptCount;
                excelRange.Cells[TestCaseTableConstants.RowIndex.TC_TEST_VIA_ANALYSIS_COUNT, dCol] = g_tTestCaseTable.dByCodeAnalysisCount;
                excelRange.Cells[TestCaseTableConstants.RowIndex.TC_TEST_VIA_OTHERS_COUNT, dCol] = g_tTestCaseTable.dByUnknownCount;

                excelRange.Cells[TestCaseTableConstants.RowIndex.TC_FUNC_COUNT, dCol] = g_tTestCaseTable.dTestCaseFuncCount;
                excelRange.Cells[TestCaseTableConstants.RowIndex.ERROR_COUNT, dCol] = g_tTestCaseTable.dNGEntryCount;

                // Hi-light the error cell.
                if (0 < g_tTestCaseTable.dNGEntryCount)
                    excelRange.Cells[TestCaseTableConstants.RowIndex.ERROR_COUNT, dCol].Interior.Color = Constants.Color.RED;
            }
            catch (System.Exception ex)
            {
                Logger.Print(sFuncName, ex.ToString());
                return false;
            }

            return true;
        }

        public int HighLightIncorrectCells(Excel.Range a_excelRange)
        {
            string sFuncName = "[HighLightIncorrectCells]";

            int i;
            int dRow;
            int dErrorCount = 0;
            int dNGCount = 0;
            TestCaseItem tTestCase;
            string sHeader;
            bool bOK;

            Logger.Print("\n" + sFuncName, "Checking & highlighting incorrect cells in the output report...");

            // Highlight incorrect cells.
            for (i = 0, dRow = TestCaseTableConstants.FIRST_ROW; i < g_tTestCaseTable.ltItems.Count; i++, dRow++)
            {
                tTestCase = g_tTestCaseTable.ltItems[i];
                sHeader = " * Row " + dRow.ToString() + ":";
                bOK = true;

                // Check source file name.
                if (tTestCase.sSourceFileName.StartsWith(Constants.StringTokens.ERROR_MSG_HEADER))
                {
                    Logger.Print(sHeader, ErrorMessage.INVLAID_SOURCE_FILE_NAME + ": \"" + tTestCase.sSourceFileName + "\"");
                    a_excelRange.Cells[dRow, TestCaseTableConstants.ColumnIndex.SOURCE_FILE].Interior.Color = Constants.Color.RED;
                    bOK = false;
                    dErrorCount++;
                }

                // Check method name.
                if (tTestCase.sMethodName.StartsWith(Constants.StringTokens.ERROR_MSG_HEADER))
                {
                    Logger.Print(sHeader, ErrorMessage.INVLAID_METHOD_NAME + ": \"" + tTestCase.sMethodName + "\"");
                    a_excelRange.Cells[dRow, TestCaseTableConstants.ColumnIndex.METHOD_NAME].Interior.Color = Constants.Color.RED;
                    bOK = false;
                    dErrorCount++;
                }

                // Check test case label name.
                if (tTestCase.sTCLabelName.StartsWith(Constants.StringTokens.ERROR_MSG_HEADER))
                {
                    Logger.Print(sHeader, ErrorMessage.INVLAID_TC_LABEL + ": \"" + tTestCase.sTCLabelName + "\"");
                    a_excelRange.Cells[dRow, TestCaseTableConstants.ColumnIndex.TC_LABEL].Interior.Color = Constants.Color.RED;
                    bOK = false;
                    dErrorCount++;
                }
                else if (tTestCase.bIsRepeated) // test case label is same as others
                {
                    Logger.Print(sHeader, ErrorMessage.DUPLICATE_TC_LABEL_FOUND + ": \"" + tTestCase.sTCLabelName + "\"");
                    a_excelRange.Cells[dRow, TestCaseTableConstants.ColumnIndex.TC_LABEL].Interior.Color = Constants.Color.RED;
                    a_excelRange.Cells[dRow, TestCaseTableConstants.ColumnIndex.NOTE].Interior.Color = Constants.Color.RED;
                    bOK = false;
                    dErrorCount++;
                }

                // Check test case function name.
                if (tTestCase.sTCFuncName.StartsWith(Constants.StringTokens.ERROR_MSG_HEADER))
                {
                    Logger.Print(sHeader, ErrorMessage.INVLAID_TC_FUNC_NAME + ": \"" + tTestCase.sTCFuncName + "\"");
                    a_excelRange.Cells[dRow, TestCaseTableConstants.ColumnIndex.TC_NAME].Interior.Color = Constants.Color.RED;
                    bOK = false;
                    dErrorCount++;
                }

                // Check test means.
                if (TestMeans.UNKNOWN == tTestCase.eTestMeans)
                {
                    Logger.Print(sHeader, ErrorMessage.TC_TEST_MEANS_SHALL_NOT_BE_UNKNOWN + ": \"" + tTestCase.sTCLabelName + "\"");
                    a_excelRange.Cells[dRow, TestCaseTableConstants.ColumnIndex.NOTE].Interior.Color = Constants.Color.RED;
                    bOK = false;
                    dErrorCount++;
                }

                // Mark the entry as NG.
                if (!bOK)
                {
                    dNGCount++;
                    a_excelRange.Cells[dRow, TestCaseTableConstants.ColumnIndex.NG_MARKER] = Constants.StringTokens.X;
                }
            }

            g_tTestCaseTable.dErrorCount = dErrorCount;
            return dNGCount;
        }

        public bool WriteDetailSheet(Excel.Workbook a_excelBook)
        {
            string sFuncName = "[WriteDetailSheet]";

            Excel.Worksheet excelSheet;
            Excel.Range excelRange;

            try
            {
                // Get the used range of the 1st sheet.
                excelSheet = (Excel.Worksheet)a_excelBook.Worksheets.get_Item(2);
                excelRange = excelSheet.UsedRange;

                // Sort test cases by source file names.
                g_tTestCaseTable.ltItems = g_tTestCaseTable.ltItems.OrderBy(x => x.sSourceFileName).ToList();

                // Write each item to EXCEL table.
                int i;
                int dRow;
                for (i = 0, dRow = TestCaseTableConstants.FIRST_ROW; i < g_tTestCaseTable.ltItems.Count; i++, dRow++)
                {
                    excelRange.Cells[dRow, TestCaseTableConstants.ColumnIndex.METHOD_NAME] = g_tTestCaseTable.ltItems[i].sMethodName;
                    excelRange.Cells[dRow, TestCaseTableConstants.ColumnIndex.SOURCE_FILE] = g_tTestCaseTable.ltItems[i].sSourceFileName;
                    excelRange.Cells[dRow, TestCaseTableConstants.ColumnIndex.TC_LABEL] = g_tTestCaseTable.ltItems[i].sTCLabelName;
                    excelRange.Cells[dRow, TestCaseTableConstants.ColumnIndex.TC_NAME] = g_tTestCaseTable.ltItems[i].sTCFuncName;
                    excelRange.Cells[dRow, TestCaseTableConstants.ColumnIndex.TDS_FILE] = g_tTestCaseTable.ltItems[i].sTDSFileName;
                    excelRange.Cells[dRow, TestCaseTableConstants.ColumnIndex.TC_SOURCE_FILE] = g_tTestCaseTable.ltItems[i].sTCSourceFileName;
                    excelRange.Cells[dRow, TestCaseTableConstants.ColumnIndex.NOTE] = g_tTestCaseTable.ltItems[i].sTCNote;
                }

                // Hi-light incorrect cells.
                int dNGCount = HighLightIncorrectCells(excelRange);

                // Write the summary info in the header.
                dRow = TestCaseTableConstants.COUNT_ROW;
                excelRange.Cells[dRow, TestCaseTableConstants.ColumnIndex.NG_MARKER] = dNGCount;
                excelRange.Cells[dRow, TestCaseTableConstants.ColumnIndex.SOURCE_FILE] = g_tTestCaseTable.dSourceFileCount;
                excelRange.Cells[dRow, TestCaseTableConstants.ColumnIndex.METHOD_NAME] = g_tTestCaseTable.dMethodCount;
                excelRange.Cells[dRow, TestCaseTableConstants.ColumnIndex.TC_LABEL] = g_tTestCaseTable.dNormalEntryCount;
                excelRange.Cells[dRow, TestCaseTableConstants.ColumnIndex.TC_NAME] = g_tTestCaseTable.dTestCaseFuncCount;

                // Filter out OK enteirs & show NG entries only (for viewing the NG entries easily).
                if (0 < dNGCount)
                {
                    try
                    {
                        excelRange.AutoFilter(TestCaseTableConstants.ColumnIndex.NG_MARKER, Constants.StringTokens.X, Excel.XlAutoFilterOperator.xlAnd, Type.Missing, true);
                    }
                    catch (SystemException e)
                    {
                        Logger.Print("", e.ToString());
                    }
                }
                g_tTestCaseTable.dNGEntryCount = dNGCount;
            }
            catch (System.Exception ex)
            {
                Logger.Print(sFuncName, ex.ToString());
                return false;
            }

            return true;
        }

        public bool SaveResults(string a_sTemplateFile, string a_sOutFile)
        {
            string sFuncName = "[SaveResults]";

            Excel.Workbook excelBook = null;

            // Check the input parameters.
            if (!File.Exists(g_sTemplateFile))
            {
                Logger.Print(sFuncName, ErrorMessage.CANNOT_FIND_TEMPLATE + ": \"" + g_sTemplateFile + "\"");
                return false;
            }
            if ("" == a_sOutFile)
            {
                Logger.Print(sFuncName, ErrorMessage.OUTPUT_FILE_IS_NULL);
                return false;
            }
            if (0 >= g_tTestCaseTable.ltItems.Count)
            {
                Logger.Print(sFuncName, ErrorMessage.NO_ENTRY_TO_BE_SAVED);
                return false;
            }

            if (null == g_excelApp)
            {
                Logger.Print(sFuncName, ErrorMessage.EXCEL_APP_IS_NULL);
                return false;
            }

            try
            {
                // Open the template file.
                excelBook = OpenExcelWorkbook(g_excelApp, a_sTemplateFile, true); // read only

                // Write the detail sheet.
                if (!WriteDetailSheet(excelBook))
                    return false;

                // Write the summary sheet. 
                // Note: This step must be behind the write-details step.
                if (!WriteSummarySheet(excelBook, a_sOutFile))
                    return false;

                // Save the modified template as the output file.
                g_excelApp.DisplayAlerts = false; // show no alert while overwritten old file
                excelBook.SaveAs(a_sOutFile);
            }
            catch (System.Exception ex)
            {
                Logger.Print(sFuncName, ex.ToString());
                return false;
            }
            finally
            {
                excelBook.Close();
            }

            return true;
        }



    }
}
