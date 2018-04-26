using System;
using System.Collections.Generic;
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

        public string PrepareSummaryReport(string output_path)
        {
            string dest = null;

            try
            {
                // get the file name of summary report
                String fileName = Path.GetFileName(g_sSummaryReport);

                // append the file name to output path
                dest = output_path + SummaryReport.SUMMARY_REPORT_FILENAME;


                //Remove old output file.
                if (File.Exists(dest))
                {
                    File.Delete(dest);
                }

                File.Copy(g_sSummaryReport, dest);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error occurred when prepare the summary report\r\n" + ex.Message);
                return null;
            }

            return dest;
        }




        //
        // Get the letter of Column in Excel.
        //
        private string GetExcelColumnName(int columnNumber)
        {
            int dividend = columnNumber;
            string columnName = String.Empty;
            int modulo;

            while (dividend > 0)
            {
                modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
                dividend = (int)((dividend - modulo) / 26);
            }

            return columnName;
        }




        //
        // Write the information of ut for each module into Summary report.
        //
        public bool WriteSummaryReport(string a_sExcelFile, ref List<ModuleInfo> items)
        {
            string sFuncName = "[WriteSummaryReport]";

            Excel.Workbook excelBook = null;
            Excel.Worksheet excelSheet;
            Excel.Range excelRange;

            if (null == a_sExcelFile)
            {
                LogToFile(sFuncName, "summary report is null.");
                return false;
            }

            // Check the readiness of the EXCEL app.
            if (null == g_excelApp)
            {
                LogToFile(sFuncName, "EXCEL app is null.");
                return false;
            }

            try
            {
                // Open the EXCEL file.
                excelBook = OpenExcelWorkbook(g_excelApp, a_sExcelFile, false);
                excelSheet = excelBook.Worksheets.get_Item(SummaryReport.SHEET_NAME);
                excelRange = excelSheet.UsedRange;

                // a count indeicates the row of starting to write the date.
                int dRawCount = SummaryReport.FIRST_ROW;

                foreach (ModuleInfo item in items)
                {
                    // index
                    excelRange.Cells[dRawCount, SummaryReport.ColumnIndex.INDEX] = dRawCount - SummaryReport.FIRST_ROW + 1;

                    // name
                    excelRange.Cells[dRawCount, SummaryReport.ColumnIndex.NAME] = item.name;

                    excelRange.Cells[dRawCount, SummaryReport.ColumnIndex.SOURCE_COUNT] = item.testCase.dSourceFileCount;
                    excelRange.Cells[dRawCount, SummaryReport.ColumnIndex.METHOD_COUNT] = item.testCase.dMethodCount;
                    excelRange.Cells[dRawCount, SummaryReport.ColumnIndex.TESTCASE_COUNT] = item.testCase.dTestCaseFuncCount;

                    // no test needed
                    excelRange.Cells[dRawCount, SummaryReport.ColumnIndex.GETTER_SETTER] = item.gettersetter;
                    excelRange.Cells[dRawCount, SummaryReport.ColumnIndex.EMPTY] = item.emptymethod;
                    excelRange.Cells[dRawCount, SummaryReport.ColumnIndex.INTERFACE] = item.interfacemethod;
                    excelRange.Cells[dRawCount, SummaryReport.ColumnIndex.ABSTRACE] = item.abstractmethod;
                    excelRange.Cells[dRawCount, SummaryReport.ColumnIndex.NATIVE] = item.nativemethod;

                    // test script
                    excelRange.Cells[dRawCount, SummaryReport.ColumnIndex.MOCKITO] = item.mockito;
                    excelRange.Cells[dRawCount, SummaryReport.ColumnIndex.POWERMOCKIT] = item.powermockito;

                    // by code analysis
                    excelRange.Cells[dRawCount, SummaryReport.ColumnIndex.BY_CODE_ANALYSIS] = item.codeanalysis;
                    excelRange.Cells[dRawCount, SummaryReport.ColumnIndex.PURE_CALL] = item.purefunctioncalls;
                    excelRange.Cells[dRawCount, SummaryReport.ColumnIndex.PURE_UI_CALL] = item.pureUIfunctioncalls;

                    // unknow
                    excelRange.Cells[dRawCount, SummaryReport.ColumnIndex.UNKNOW] = item.unknow;


                    excelRange.Cells[dRawCount, SummaryReport.ColumnIndex.TOTAL_TESTCASE_COUNT] = item.count;
                    excelRange.Cells[dRawCount, SummaryReport.ColumnIndex.NORMAL_ENTRY] = item.testCase.dNormalEntryCount;
                    excelRange.Cells[dRawCount, SummaryReport.ColumnIndex.REPEATED_ENTRY] = item.testCase.dRepeatedEntryCount;
                    excelRange.Cells[dRawCount, SummaryReport.ColumnIndex.ERROR_ENTRY] = item.testCase.dErrorEntryCount;

                    excelRange.Cells[dRawCount, SummaryReport.ColumnIndex.NG_COUNT] = item.testCase.dNGEntryCount;

                    excelRange.Cells[dRawCount, SummaryReport.ColumnIndex.ERROR_COUNT] = item.testCase.dErrorCount;

                    if (item.testCase.dErrorCount > 0)
                    {
                        excelRange.Cells[dRawCount, SummaryReport.ColumnIndex.ERROR_COUNT].Interior.Color = Constants.Color.RED;
                    }


                    dRawCount++;
                }

                // Save the modified template as the output file.
                g_excelApp.DisplayAlerts = false; // show no alert while overwritten old file
                excelBook.Save();
            }
            catch (System.Exception ex)
            {
                LogToFile(sFuncName, Path.GetFileName(a_sExcelFile) + ": " + ex.ToString());
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
