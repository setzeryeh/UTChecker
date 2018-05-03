using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace UTChecker
{
    public partial class UTChecker
    {

        private List<string> SearchTDSFiles(string a_sStartPath, string a_sOutFile)
        {
            string sFuncName = "[SearchTDSFiles]";

            List<string> a_lsOutList = new List<string>();

            // Check the input parameters.
            if ("" == a_sStartPath)
            {
                Logger.Print(sFuncName, "Null start path is specified.");
                return null;
            }

            // reset
            a_lsOutList.Clear();

            // Serach and collect all log files recursively.
            CollectFiles(a_sStartPath, Constants.TDS_FILENAME_EXT, Constants.TDS_FILENAME_PREFIX, ref a_lsOutList);

            // Save the list of found files to the specifed file.
            if ("" != a_sOutFile)
            {
                WriteStringListToTextFile(ref a_lsOutList, a_sOutFile);
            }

            Logger.Print(sFuncName, a_lsOutList.Count.ToString() + " TDS file(s) collected.");


            return a_lsOutList;
        }





        /// <summary>
        /// 
        /// </summary>
        /// <param name="a_sDir"></param>
        /// <param name="a_sFileExt"></param>
        /// <param name="a_sToken"></param>
        /// <param name="a_lsCollection"></param>
        /// <returns></returns>
        private List<string> CollectFiles(string a_sDir, string a_sFileExt, string a_sToken, ref List<string> a_lsCollection)
        {
            string sFuncName = "[CollectFiles]";
            string sFileName;

            try
            {
                // Check the existence of the specified path.
                if (!Directory.Exists(a_sDir))
                {
                    Logger.Print(sFuncName, "Cannot find path \"" + a_sDir + "\"; skipped.");
                    return a_lsCollection;
                }

                // Collect the considered files stored in current folder.
                string[] FileList = Directory.GetFiles(a_sDir, a_sFileExt);
                foreach (string f in FileList)
                {
                    // Discard the path from the name.
                    sFileName = Path.GetFileName(f);

                    // Check if the file name starts with the spcified token.
                    // If yes, add it in the list.
                    if (sFileName.StartsWith(a_sToken))
                    {
                        a_lsCollection.Add(f);
                    }
                }

                // Collect the considered files stored in sub-folders.
                string[] DirList = Directory.GetDirectories(a_sDir);
                foreach (string d in DirList)
                {
                    a_lsCollection = CollectFiles(d, a_sFileExt, a_sToken, ref a_lsCollection);
                }
            }
            catch (System.Exception excpt)
            {
                Logger.Print(sFuncName, excpt.Message);
            }

            return a_lsCollection;
        }


        /// <summary>
        /// 
        /// </summary>
        /// <param name="a_lsInList"></param>
        /// <param name="a_sOutFile"></param>
        /// <returns></returns>
        private bool WriteStringListToTextFile(ref List<string> a_lsInList, string a_sOutFile)
        {
            string sFuncName = "[WriteStringListToTextFile]";

            // Check the input.
            if (null == a_lsInList)
            {
                Logger.Print(sFuncName, "Cannot save a null list to file.");
                return false;
            }
            if ("" == a_sOutFile)
            {
                Logger.Print(sFuncName, "No output file is specified.");
                return false;
            }

            // Check the number of lines to be saved.
            if (0 == a_lsInList.Count)
            {
                Logger.Print(sFuncName + "The list to be saved is an empty list. Do nothing.");
                return true;
            }

            // Write the error log to the output file.
            try
            {
                using (StreamWriter sw = File.AppendText(a_sOutFile))
                {
                    foreach (string sLine in a_lsInList)
                    {
                        sw.WriteLine(sLine);
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Print(sFuncName, ex.ToString());
                return false;
            }

            return true;
        }


        /// <summary>
        /// 
        /// </summary>
        /// <param name="a_sStartPath"></param>
        /// <param name="a_lsOutList"></param>
        /// <param name="a_sOutFile"></param>
        /// <returns></returns>
        private List<string> SearchMockitoTestLogs(string a_sStartPath, string a_sOutFile)
        {
            string sFuncName = "[SearchTestLogs]";
            string dir_JUnit_Auto = "JUnit_Auto";
            string dir_JUnit_Local = "JUnit_Local";

            List<string> a_lsOutList = new List<string>();


            // Check the input parameters.
            if ("" == a_sStartPath)
            {
                Logger.Print(sFuncName, "Null start path is specified.");
                return null;
            }

            // clear 
            a_lsOutList.Clear();

            a_sStartPath = a_sStartPath + "\\" + dir_JUnit_Auto + "\\";

            // Serach and collect all log files recursively.
            CollectFiles(a_sStartPath, Constants.TESTLOG_FILENAME_EXT, "", ref a_lsOutList);

            // Save the list of found files to the specifed file.
            if ("" != a_sOutFile)
            {
                WriteStringListToTextFile(ref a_lsOutList, a_sOutFile);
            }

            Logger.Print(sFuncName, a_lsOutList.Count.ToString() + " Test Log(s) collected.");

            return a_lsOutList;
        }


        /// <summary>
        /// 
        /// </summary>
        /// <param name="a_sStartPath"></param>
        /// <param name="a_lsOutList"></param>
        /// <param name="a_sOutFile"></param>
        /// <returns></returns>
        private List<string> SearchPowerMockitoTestLogs(string a_sName, string a_sLogPath)
        {
            string sFuncName = "[SearchTestLogs]";
            string dir_PowerMockito = "PowerMockito";

            string a_sStartPath = a_sLogPath + a_sName + "\\";

            //string sTempListForLog = a_sStartPath + a_sName + "_testlogs.list";
            string sTempListFileName = a_sName + "_testlogs.txt";


            List<string> a_lsOutList = new List<string>();

            // Check the input parameters.
            if ("" == a_sStartPath)
            {
                Logger.Print(sFuncName, "Null start path is specified.");
                return null;
            }

            if (File.Exists(sTempListFileName))
            {
                File.Delete(sTempListFileName);
            }


            // clear 
            a_lsOutList.Clear();

            // replace the path to find test logs
            a_sStartPath = a_sStartPath + dir_PowerMockito + "\\";

            // Serach and collect all log files recursively.
            CollectFiles(a_sStartPath, Constants.TESTLOG_FILENAME_EXT, "", ref a_lsOutList);

            // Save the list of found files to the specifed file.
            if ("" != sTempListFileName)
            {
                //WriteStringListToTextFile(ref a_lsOutList, sTempListForLog);
            }

            List<string> logs = processLogs(a_lsOutList, sTempListFileName);

            Logger.Print(sFuncName, a_lsOutList.Count.ToString() + " Test Log(s) collected.");


            return logs;
        }


        /// <summary>
        /// 
        /// </summary>
        /// <param name="logPath"></param>
        /// <param name="tempFile"></param>
        /// <returns></returns>
        private List<string> processLogs(List<string> logPath, string tempFile)
        {
            List<string> logSet = new List<string>();
            logSet.Clear();

            // remove old for create new one
            if (File.Exists(tempFile))
            {
                File.Delete(tempFile);
            }

            
            foreach (string fullName in logPath)
            {
                // get the file name of test log.
                string logFileName = Path.GetFileName(fullName);

                // get the full path of log file
                string fullPath = Path.GetDirectoryName(fullName);

                // get the class from previous direcotry.
                string parent = Directory.GetParent(fullPath).FullName;
                string className = fullPath.Substring(parent.Length + 1); // with '\' plus 1

                string data = $"{className}, {logFileName}, {fullName}";
                logSet.Add(data);

                //using (StreamWriter sw = File.AppendText(tempFile))
                //{
                //    sw.WriteLine(data);
                //}

            }

            return logSet;
        }


    }
}
