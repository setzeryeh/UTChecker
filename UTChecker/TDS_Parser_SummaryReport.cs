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

        /// <summary>
        /// 
        /// </summary>
        /// <param name="a_sSummaryReportFile"></param>
        /// <returns></returns>
        public List<string> ReadAllModuleNames(string a_sSummaryReportFile)
        {
            string sFuncName = "[ReadSummaryModuleTable]";

            List<string> lsModuleNames = null;

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
                excelBook = g_excelApp.Workbooks.Open(a_sSummaryReportFile, 0, true, 6, "", "", true,
                    Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);

                // Get the used range of the 1st sheet.
                excelSheet = (Excel.Worksheet)excelBook.Worksheets.get_Item(SummaryReport.SHEET_NAME);
                excelRange = excelSheet.UsedRange;

                lsModuleNames = new List<string>();

                for (int i = SummaryReport.FIRST_ROW; i <= excelRange.Rows.Count; i++)
                {
                    // Read the module names (1st row is caption). 
                    string sModuleName = ReadStringFromExcelCell(excelRange.Cells[i, SummaryReport.ColumnIndex.MODULE_NAME], "", false);
                    if ("" == sModuleName)
                    {
                        Logger.Print(sFuncName, "Row " + i.ToString() + ": Module name is empty.");
                    }


                    lsModuleNames.Add(sModuleName);
                }

            }
            catch (Exception ex)
            {
                Logger.Print(sFuncName, ex.ToString());

                return null;

            }
            finally
            {
                // Close the EXCEL table.
                if (null != excelBook)
                {
                    excelBook.Close();
                }
            }

            return lsModuleNames;
        }


        /// <summary>
        /// 
        /// </summary>
        /// <param name="a_sModuleName"></param>
        /// <param name="lsModuleNames"></param>
        /// <returns></returns>
        public static int GetModuleId(string a_sModuleName, List<string> lsModuleNames)
        {
            a_sModuleName = a_sModuleName.Replace("_", " ");

            for (int i = 0; i < lsModuleNames.Count; i++)
            {
                if (a_sModuleName == lsModuleNames[i])
                    return i+1;
            }

            return -1;
        }







        public string PrepareSummaryReport(string output_path)
        {
            string dest = null;

            try
            {
                // get the file name of summary report
                String fileName = Path.GetFileName(g_sSummaryReport);

                // append the file name to output path
                dest = output_path + SummaryReport.FILE_NAME;


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
                Logger.Print(sFuncName, "summary report is null.");
                return false;
            }

            // Check the readiness of the EXCEL app.
            if (null == g_excelApp)
            {
                Logger.Print(sFuncName, "EXCEL app is null.");
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
                    excelRange.Cells[dRawCount, SummaryReport.ColumnIndex.MODULE_NAME] = item.name;

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
                Logger.Print(sFuncName, Path.GetFileName(a_sExcelFile) + ": " + ex.ToString());
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
