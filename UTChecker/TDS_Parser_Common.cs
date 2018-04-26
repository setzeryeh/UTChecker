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

        private bool ReadInfoFromTDSCoverSheet(Excel.Workbook a_excelBook, ref string a_sSourceFileName, ref string a_sMethodName)
        {
            Excel.Worksheet excelSheet;
            Excel.Range excelRange;
            string sSourceFileName;
            string sValue;
            int dRow, dCol;
            const string sSheetName = "Coversheet";
            bool bIsJava, bIsCOrCpp;


            // -------------------------------------------------------------------------
            // Read data form the "Coversheet" sheet.
            // -------------------------------------------------------------------------
            try
            {
                excelSheet = (Excel.Worksheet)a_excelBook.Worksheets.get_Item(sSheetName);
                excelRange = excelSheet.UsedRange;
            }
            catch
            {
                LogToFile(Constants.StringTokens.MSG_BULLET, "No \"" + sSheetName + "\" sheet can be found.");
                return false;
            }

            // Locate the considered cell (where y in [2,4] and x in [2,3]).
            bool bFound = false;
            dCol = 2;
            for (dRow = 2; dRow <= 3; dRow++)
            {
                for (dCol = 2; dCol <= 4; dCol++)
                {
                    sValue = ReadStringFromExcelCell(excelRange.Cells[dRow, dCol], "", true);
                    if ("File" == sValue)
                    {
                        bFound = true;
                        break;
                    }
                }
                if (bFound)
                    break;
            }

            if (bFound)
            {
                // Read the file name & SVN revision.
                sSourceFileName = ReadStringFromExcelCell(excelRange.Cells[dRow, dCol + 1], "", true);

                // Check if the source file name contains any space, and remove it.
                // e.g. "XXX .java" --> "XXX.java"
                if (sSourceFileName.Contains(" "))
                {
                    LogToFile(Constants.StringTokens.MSG_BULLET, "File name \"" + sSourceFileName + "\" contains space(s). Stripped.");
                    sSourceFileName.Replace(" ", "");
                }

                // Check if the string read is a java/c/cpp file.
                string sTmp = sSourceFileName.ToLower();
                bIsJava = sTmp.EndsWith(".java");
                bIsCOrCpp = (sTmp.EndsWith(".c") || sTmp.EndsWith(".cpp"));

                // Set the source file name read and increase the row.
                if (bIsCOrCpp || bIsJava)
                {
                    a_sSourceFileName = sSourceFileName;
                    dRow++;
                }
            }
            else
            {
                LogToFile(Constants.StringTokens.MSG_BULLET, "No source file name can be found from the \"" + sSheetName + "\" sheet.");
                return false;
            }

            // Extrace method name.
            a_sMethodName = "";
            if (bIsCOrCpp)
            {
                // Locate the method name cell.
                bFound = false;
                dRow++;
                for (; dRow <= 7; dRow++)
                {
                    for (dCol = 1; dCol <= 3; dCol++)
                    {
                        sValue = ReadStringFromExcelCell(excelRange.Cells[dRow, dCol], "", true);
                        if ("Method Name" == sValue)
                        {
                            bFound = true;
                            break;
                        }
                    }

                    // Extract method name.
                    if (bFound)
                    {
                        dRow++;
                        dCol = dCol + 2;
                        sValue = ReadStringFromExcelCell(excelRange.Cells[dRow, dCol], "", true);
                        a_sMethodName = ExtractMethodName(sValue);

                        return ("" != a_sMethodName);
                    }
                }

                return false;
            }

            return true;
        }

        private bool ReadDataFromTDSFiles(string a_sModuleName, ref List<string> a_lsTDSFiles)
        {
            string sFuncName = "[ReadDataFromTDSFiles]";

            Excel.Workbook excelBook = null;
            bool bIsJava;
            string sFileNameWithoutPath;
            string sShortTDSFileName;
            string sSourceFileName = "";
            string sMethodName = "";
            int dErrorCount = 1;
            int dProceedFileCount = 0;


            // Check the input list.
            if (null == a_lsTDSFiles)
            {
                LogToFile(sFuncName, "The input list is null.");
                return false;
            }
            if (0 == a_lsTDSFiles.Count)
            {
                LogToFile(sFuncName, "No TDS file is found.");
                return false;
            }
            // Check the EXCEL app.
            if (null == g_excelApp)
            {
                LogToFile(sFuncName, "EXCEL app is null.");
                return false;
            }

            try
            {
                LogToFile(sFuncName, "Reading TDS files...");

                // Initialize objects.
                dErrorCount = 0;
                g_tTestCaseTable.ltItems.Clear();
                g_excelApp.DisplayAlerts = false; // show no alert while closing the file

                // Read data from each TDS file.
                foreach (string sFile in a_lsTDSFiles)
                {
                    sFileNameWithoutPath = "\"" + Path.GetFileName(sFile) + "\"";
                    LogToFile("", "Reading " + sFileNameWithoutPath + "...");

                    // Check the existence of the TDS file.
                    if (!File.Exists(sFile))
                    {
                        LogToFile(Constants.StringTokens.MSG_BULLET, "Cannot find " + sFile);
                        dErrorCount++;
                        continue;
                    }

                    // Open the TDS file & get the lookup-table sheet.
                    excelBook = OpenExcelWorkbook(g_excelApp, sFile, true);  // ready only
                    if (null == excelBook)
                        continue;

                    try
                    {
                        // Read source file name from the cover sheet.
                        if (!ReadInfoFromTDSCoverSheet(excelBook, ref sSourceFileName, ref sMethodName))
                        {
                            dErrorCount++;
                            continue;
                        }

                        // Determine the source file type.
                        bIsJava = (sSourceFileName.EndsWith("java"));

                        // Read data form the "TestCase" sheet.
                        sShortTDSFileName = sFile.Replace(g_sTDSPath + a_sModuleName + "\\", "");
                        if (bIsJava)
                            ReadTestCasesFromTDSFile_Java(excelBook, ref sShortTDSFileName, ref sSourceFileName);
                        else
                            ReadTestCasesFromTDSFile_C(excelBook, ref sShortTDSFileName, ref sSourceFileName, ref sMethodName);

                        dProceedFileCount++;
                    }
                    catch (SystemException ex)
                    {
                        LogToFile(sFuncName, sFile.Replace(g_sTDSPath, "...") + ": " + ex.ToString());
                        dErrorCount++;
                    }
                    finally
                    {
                        // Close the TDS file.
                        excelBook.Close();
                    }
                }
            }
            catch (SystemException ex)
            {
                LogToFile(sFuncName, ex.ToString());
                dErrorCount++;
            }

            // Show the # of proceeded files.
            if (dProceedFileCount != a_lsTDSFiles.Count)
                LogToFile(sFuncName, dProceedFileCount.ToString() + " of " + a_lsTDSFiles.Count + " TDS files proceed.");

            return true;
        }


    }
}
