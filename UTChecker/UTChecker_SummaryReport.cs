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

                File.Copy(g_sSummaryReport, dest);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error occurred when prepare the summary report\r\n" + ex.Message);
                return null;
            }

            return dest;
        }




        /// <summary>
        /// 
        /// </summary>
        /// <param name="a_sSummaryReportFile"></param>
        /// <returns></returns>
        public List<string> ReadAllModuleNamesFromExcel(string a_sSummaryReportFile)
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
        public int GetModuleId(string a_sModuleName, List<string> lsModuleNames)
        {
            a_sModuleName = a_sModuleName.Replace("_", " ");

            for (int i = 0; i < lsModuleNames.Count; i++)
            {
                if (a_sModuleName == lsModuleNames[i])
                    return i + 1;
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
        public bool WriteSummaryReport(string a_sExcelFile, ModuleInfo item, int index)
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
                
                
                // index
                //excelRange.Cells[int, SummaryReport.ColumnIndex.INDEX] = dRawCount - SummaryReport.FIRST_ROW + 1;

                // name
                excelRange.Cells[index, SummaryReport.ColumnIndex.MODULE_NAME] = item.name;

                excelRange.Cells[index, SummaryReport.ColumnIndex.SOURCE_COUNT] = item.testCase.dSourceFileCount;
                excelRange.Cells[index, SummaryReport.ColumnIndex.METHOD_COUNT] = item.testCase.dMethodCount;
                excelRange.Cells[index, SummaryReport.ColumnIndex.TESTCASE_COUNT] = item.testCase.dTestCaseFuncCount;

                // no test needed
                excelRange.Cells[index, SummaryReport.ColumnIndex.GETTER_SETTER] = item.gettersetter;
                excelRange.Cells[index, SummaryReport.ColumnIndex.EMPTY] = item.emptymethod;
                excelRange.Cells[index, SummaryReport.ColumnIndex.INTERFACE] = item.interfacemethod;
                excelRange.Cells[index, SummaryReport.ColumnIndex.ABSTRACE] = item.abstractmethod;
                excelRange.Cells[index, SummaryReport.ColumnIndex.NATIVE] = item.nativemethod;

                // test script
                excelRange.Cells[index, SummaryReport.ColumnIndex.MOCKITO] = item.mockito;
                excelRange.Cells[index, SummaryReport.ColumnIndex.POWERMOCKIT] = item.powermockito;

                // by code analysis
                excelRange.Cells[index, SummaryReport.ColumnIndex.BY_CODE_ANALYSIS] = item.codeanalysis;
                excelRange.Cells[index, SummaryReport.ColumnIndex.PURE_CALL] = item.purefunctioncalls;
                excelRange.Cells[index, SummaryReport.ColumnIndex.PURE_UI_CALL] = item.pureUIfunctioncalls;

                // unknow
                excelRange.Cells[index, SummaryReport.ColumnIndex.UNKNOW] = item.unknow;


                excelRange.Cells[index, SummaryReport.ColumnIndex.TOTAL_TESTCASE_COUNT] = item.count;
                excelRange.Cells[index, SummaryReport.ColumnIndex.NORMAL_ENTRY] = item.testCase.dNormalEntryCount;
                excelRange.Cells[index, SummaryReport.ColumnIndex.REPEATED_ENTRY] = item.testCase.dRepeatedEntryCount;
                excelRange.Cells[index, SummaryReport.ColumnIndex.ERROR_ENTRY] = item.testCase.dErrorEntryCount;

                excelRange.Cells[index, SummaryReport.ColumnIndex.NG_COUNT] = item.testCase.dNGEntryCount;

                excelRange.Cells[index, SummaryReport.ColumnIndex.ERROR_COUNT] = item.testCase.dErrorCount;

                if (item.testCase.dErrorCount > 0)
                {
                    excelRange.Cells[index, SummaryReport.ColumnIndex.ERROR_COUNT].Interior.Color = Constants.Color.RED;
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
