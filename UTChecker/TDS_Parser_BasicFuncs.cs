using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace UTChecker
{
    public partial class TDS_Parser
    {

        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        private bool UpdateUTCheckerSettingToFile()
        {

            if (File.Exists(UTCheckerSetting.FileName))
            {
                File.Delete(UTCheckerSetting.FileName);
            }

            using (System.IO.StreamWriter file = new System.IO.StreamWriter(UTCheckerSetting.FileName))
            {

                file.WriteLine(UTCheckerSetting.Prefix + " " +
                                UTCheckerSetting.ListFile + "=" +
                                g_FilePathSetting.listFile);

                file.WriteLine(UTCheckerSetting.Prefix + " " +
                                UTCheckerSetting.TDSPath + "=" +
                                g_FilePathSetting.tdsPath);

                file.WriteLine(UTCheckerSetting.Prefix + " " +
                                UTCheckerSetting.OutputPath + "=" +
                                g_FilePathSetting.outputPath);


                file.WriteLine(UTCheckerSetting.Prefix + " " +
                                UTCheckerSetting.ReportTemplate + "=" +
                                g_FilePathSetting.reportTemplate);


                file.WriteLine(UTCheckerSetting.Prefix + " " +
                                UTCheckerSetting.SummaryTemplate + "=" +
                                g_FilePathSetting.summaryTemplate);


                file.WriteLine(UTCheckerSetting.Prefix + " " +
                                UTCheckerSetting.TestLogs + "=" +
                                g_FilePathSetting.testlogsPath);
            }


            return true;
        }



        public event EventHandler UpdatePathEvent;



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






    }
}
