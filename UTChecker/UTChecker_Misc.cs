using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;


namespace UTChecker
{
    public partial class UTChecker
    {


        /// <summary>
        /// Release Office app.
        /// </summary>
        private void ReleaseOfficeApps()
        {
            string sFuncName = "[ReleaseExcelApps]";

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


                // Close Word app;
                if (null != g_wordApp)
                {
                    g_wordApp.DisplayAlerts = Microsoft.Office.Interop.Word.WdAlertLevel.wdAlertsNone;
                    g_wordApp.Quit();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(g_wordApp); // force the app to be closed
                    g_wordApp = null;
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
        private Word.Document OpenWordDocument(Word.Application a_wordApp, string a_sWordFile)
        {
            
            string sFuncName = "[OpenWordDocument]";

            Word.Document wordDoc = null;


            if (null == a_wordApp)
            {
                Logger.Print(sFuncName, "Null Word app is given.");
                return null;
            }

            try
            {
                wordDoc = a_wordApp.Documents.Open(
                    a_sWordFile,    // FileName
                    false,          // ConfirmConversions
                    true,           // ReadOnly
                    Type.Missing,   // AddToRecentFiles
                    Type.Missing,   // PasswordDocument
                    Type.Missing,   // PasswordTemplate
                    Type.Missing,   // Revert
                    Type.Missing,   // WritePasswordDocument
                    Type.Missing,   // WritePasswordTempalte
                    Type.Missing,   // Format
                    Type.Missing,   // Encoding
                    false,          // Visible
                    Type.Missing,   // OpenAndRepair
                    Type.Missing,   // DocumentDirection
                    Type.Missing,   // NoEncodingDialog
                    Type.Missing);  // XMLTransform

 
            }
            catch (System.Exception ex)
            {
                Logger.Print(sFuncName, "Open Word file failed: " + a_sWordFile + " (" + ex.ToString() + ")");
                return null;
            }

            return wordDoc;

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
                Excel.Workbook excelBook = a_excelApp.Workbooks.Open(a_sExcelFile, 0, a_bReadOnly, 5, "", "",
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







        public class MessageFilter : IOleMessageFilter
        {
            //
            // Class containing the IOleMessageFilter
            // thread error-handling functions.

            // Start the filter.
            public static void Register()
            {
                IOleMessageFilter newFilter = new MessageFilter();
                IOleMessageFilter oldFilter = null;
                CoRegisterMessageFilter(newFilter, out oldFilter);
            }

            // Done with the filter, close it.
            public static void Revoke()
            {
                IOleMessageFilter oldFilter = null;
                CoRegisterMessageFilter(null, out oldFilter);
            }

            //
            // IOleMessageFilter functions.
            // Handle incoming thread requests.
            int IOleMessageFilter.HandleInComingCall(int dwCallType,
              System.IntPtr hTaskCaller, int dwTickCount, System.IntPtr
              lpInterfaceInfo)
            {
                //Return the flag SERVERCALL_ISHANDLED.
                return 0;
            }

            // Thread call was rejected, so try again.
            int IOleMessageFilter.RetryRejectedCall(System.IntPtr
              hTaskCallee, int dwTickCount, int dwRejectType)
            {
                if (dwRejectType == 2)
                // flag = SERVERCALL_RETRYLATER.
                {
                    // Retry the thread call immediately if return >=0 & 
                    // <100.
                    return 99;
                }
                // Too busy; cancel call.
                return -1;
            }

            int IOleMessageFilter.MessagePending(System.IntPtr hTaskCallee,
              int dwTickCount, int dwPendingType)
            {
                //Return the flag PENDINGMSG_WAITDEFPROCESS.
                return 2;
            }

            // Implement the IOleMessageFilter interface.
            [DllImport("Ole32.dll")]
            private static extern int
              CoRegisterMessageFilter(IOleMessageFilter newFilter, out
              IOleMessageFilter oldFilter);
        }

        [ComImport(), Guid("00000016-0000-0000-C000-000000000046"),
        InterfaceTypeAttribute(ComInterfaceType.InterfaceIsIUnknown)]
        interface IOleMessageFilter
        {
            [PreserveSig]
            int HandleInComingCall(
                int dwCallType,
                IntPtr hTaskCaller,
                int dwTickCount,
                IntPtr lpInterfaceInfo);

            [PreserveSig]
            int RetryRejectedCall(
                IntPtr hTaskCallee,
                int dwTickCount,
                int dwRejectType);

            [PreserveSig]
            int MessagePending(
                IntPtr hTaskCallee,
                int dwTickCount,
                int dwPendingType);
        }
    }

}
