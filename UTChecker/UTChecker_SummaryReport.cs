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
    public partial class UTChecker
    {



        /// <summary>
        /// 
        /// </summary>
        /// <param name="output_path"></param>
        /// <returns></returns>
        public string PrepareSummaryReport(string templatePath, string outputPath)
        {
            string sFuncName = "[PrepareSummaryReport]";

            string dest = null;

            try
            {
                // get the file name of summary report
                String fileName = Path.GetFileName(templatePath);

                // append the file name to output path
                dest = outputPath + SummaryReport.FILE_NAME;


                //Remove old output file.
                if (File.Exists(dest))
                {
                    File.Delete(dest);
                }

                File.Copy(templatePath, dest);

            }
            catch (Exception ex)
            {
                Logger.Print(sFuncName, "Error occurred when prepare the summary report. " + ex.Message);
                return null;
            }

            return dest;
        }




        /// <summary>
        /// 
        /// </summary>
        /// <param name="a_sSummaryReportFile"></param>
        /// <returns></returns>
        public Dictionary<string, int> ReadAllModuleNamesFromExcel(string a_sSummaryReportFile)
        {
            string sFuncName = "[ReadSummaryModuleTable]";

            Dictionary<string, int> lsModuleNames = null;

            Excel.Workbook excelBook = null;
            Excel.Worksheet excelSheet;
            Excel.Range excelRange;

            // Check the readiness of the Excel app.
            if (null == g_excelApp)
            {
                Logger.Print(sFuncName, "Cannot execute MS Excel.");

                return null;
            }


            // Check the existence of the summary report.
            if (!File.Exists(a_sSummaryReportFile))
            {
                Logger.Print(sFuncName, "Cannot find " + a_sSummaryReportFile);

                return null;
            }

            try
            {
                

                // Open the specified EXCEL file.
                excelBook = g_excelApp.Workbooks.Open(a_sSummaryReportFile, 0, true, 5, "", "", true,
                    Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);

                // Get the used range of the 1st sheet.
                excelSheet = (Excel.Worksheet)excelBook.Worksheets.get_Item(SummaryReport.SHEET_NAME);
                excelRange = excelSheet.UsedRange;

                lsModuleNames = new Dictionary<string, int>();

                for (int i = SummaryReport.FIRST_ROW; i <= excelRange.Rows.Count; i++)
                {
                    // Read the module names (1st row is caption). 
                    string sModuleName = ReadStringFromExcelCell(excelRange.Cells[i, SummaryReport.ColumnIndex.MODULE_NAME], "", false);
                    if ("" == sModuleName)
                    {
                        Logger.Print(sFuncName, "Row " + i.ToString() + ": Module name is empty.");
                    }

                    lsModuleNames.Add(sModuleName, i);
                }

            }
            catch (Exception ex)
            {
                Logger.Print(sFuncName, ex.ToString());

                return null;

            }
            finally
            {
                excelBook.Close(false, Type.Missing, Type.Missing);
            }

            return lsModuleNames;
        }




        /// <summary>
        /// 
        /// </summary>
        /// <param name="a_sModuleName"></param>
        /// <param name="lsModuleNames"></param>
        /// <returns></returns>
        public int GetModuleId(string a_sModuleName, List<string> lsModuleNames)
        {
            string moduleName = a_sModuleName.Replace("_", " ");

            for (int i = 0; i < lsModuleNames.Count; i++)
            {
                if (moduleName == lsModuleNames[i])
                {
                    return i + 1;
                }
            }

            return -1;
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
        public bool WriteSummaryReport(string a_sExcelFile, TestCaseTable item, int index)
        {
            string sFuncName = "[WriteSummaryReport]";

            Excel.Workbook excelBook = null;
            Excel.Worksheet excelSheet;
            Excel.Range excelRange;

            if (null == a_sExcelFile)
            {
                Logger.Print(sFuncName, "summary report is null.");
                return false;
            }

            // Check the readiness of the EXCEL app.
            if (null == g_excelApp)
            {
                Logger.Print(sFuncName, ErrorMessage.EXCEL_APP_IS_NULL);
                return false;
            }


            try
            {
                
                excelBook = g_excelApp.Workbooks.Open(a_sExcelFile, 0, false, 5, "", "", true,
                    Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);

                // Open the EXCEL file.
                //excelBook = OpenExcelWorkbook(g_excelApp, a_sExcelFile, true);


                excelSheet = excelBook.Worksheets.get_Item(SummaryReport.SHEET_NAME);
                excelRange = excelSheet.UsedRange;


                // index
                //excelRange.Cells[int, SummaryReport.ColumnIndex.INDEX] = dRawCount - SummaryReport.FIRST_ROW + 1;
                int dRow = index;
                // name
                //excelRange.Cells[dRow, SummaryReport.ColumnIndex.MODULE_NAME] = item.name;

                excelRange.Cells[dRow, SummaryReport.ColumnIndex.SOURCE_COUNT] = item.dSourceFileCount;
                excelRange.Cells[dRow, SummaryReport.ColumnIndex.METHOD_COUNT] = item.dMethodCount;
                excelRange.Cells[dRow, SummaryReport.ColumnIndex.TESTCASE_COUNT] = item.dTestCaseFuncCount;

                // test script
                excelRange.Cells[dRow, SummaryReport.ColumnIndex.MOCKITO] = item.stTestTypeStatistic.mockito;
                excelRange.Cells[dRow, SummaryReport.ColumnIndex.POWERMOCKIT] = item.stTestTypeStatistic.powermockito;

                // unknow
                excelRange.Cells[dRow, SummaryReport.ColumnIndex.VECTORCAST] = item.stTestTypeStatistic.vectorcast;

                // no test needed
                excelRange.Cells[dRow, SummaryReport.ColumnIndex.GETTER_SETTER] = item.stTestTypeStatistic.gettersetter;
                excelRange.Cells[dRow, SummaryReport.ColumnIndex.EMPTY] = item.stTestTypeStatistic.emptymethod;
                excelRange.Cells[dRow, SummaryReport.ColumnIndex.INTERFACE] = item.stTestTypeStatistic.interfacemethod;
                excelRange.Cells[dRow, SummaryReport.ColumnIndex.ABSTRACE] = item.stTestTypeStatistic.abstractmethod;
                excelRange.Cells[dRow, SummaryReport.ColumnIndex.NATIVE] = item.stTestTypeStatistic.nativemethod;


                // by code analysis
                excelRange.Cells[dRow, SummaryReport.ColumnIndex.BY_CODE_ANALYSIS] = item.stTestTypeStatistic.codeanalysis;
                excelRange.Cells[dRow, SummaryReport.ColumnIndex.PURE_CALL] = item.stTestTypeStatistic.purefunctioncalls;
                excelRange.Cells[dRow, SummaryReport.ColumnIndex.PURE_UI_CALL] = item.stTestTypeStatistic.pureUIfunctioncalls;

                // unknow
                excelRange.Cells[dRow, SummaryReport.ColumnIndex.UNKNOW] = item.stTestTypeStatistic.unknow;


                excelRange.Cells[dRow, SummaryReport.ColumnIndex.TOTAL_TESTCASE_COUNT] = item.ltItems.Count;
                excelRange.Cells[dRow, SummaryReport.ColumnIndex.NORMAL_ENTRY] = item.dNormalEntryCount;
                excelRange.Cells[dRow, SummaryReport.ColumnIndex.REPEATED_ENTRY] = item.dRepeatedEntryCount;
                excelRange.Cells[dRow, SummaryReport.ColumnIndex.ERROR_ENTRY] = item.dErrorEntryCount;
                excelRange.Cells[dRow, SummaryReport.ColumnIndex.ERROR_COUNT] = item.dErrorCount;
                excelRange.Cells[dRow, SummaryReport.ColumnIndex.NG_COUNT] = item.dNGEntryCount;
                excelRange.Cells[dRow, SummaryReport.ColumnIndex.TESTLOG_ISSUE_COUNT] = item.dTestLogIssueCount;
                excelRange.Cells[dRow, SummaryReport.ColumnIndex.SUTS_ISSUE_COUNT] = item.dSUTSIssueCount;

                // set The color of background for this cell to indicate that this is a Error.
                if (item.dNGEntryCount > 0)
                {
                    excelRange.Cells[dRow, SummaryReport.ColumnIndex.NG_COUNT].Interior.Color = Constants.Color.RED;
                }

                if (item.dTestLogIssueCount > 0)
                {
                    excelRange.Cells[dRow, SummaryReport.ColumnIndex.TESTLOG_ISSUE_COUNT].Interior.Color = Constants.Color.RED;
                }

                if (item.dSUTSIssueCount > 0)
                {
                    excelRange.Cells[dRow, SummaryReport.ColumnIndex.SUTS_ISSUE_COUNT].Interior.Color = Constants.Color.RED;
                }



                // Save the modified template as the output file.
                g_excelApp.DisplayAlerts = false; // show no alert while overwritten old file
                excelBook.Save();
            }
            catch (System.Exception ex)
            {
                Logger.Print(sFuncName, Path.GetFileName(a_sExcelFile) + ": " + ex.ToString());
                return false;
            }
            finally
            {
                excelBook.Close(false, Type.Missing, Type.Missing);
            }

            return true;
        }


    }
}
