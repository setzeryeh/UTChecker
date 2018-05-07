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

        /// <summary>
        /// 
        /// </summary>
        /// <param name="a_sStartPath"></param>
        /// <param name="a_sOutFile"></param>
        /// <returns></returns>
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
#if DEBUG
                WriteStringListToTextFile(ref a_lsOutList, a_sOutFile);
#endif
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
        /// <param name="a_sName"></param>
        /// <param name="a_sLogPath"></param>
        /// <returns></returns>
        public List<TestLog> CollectTestLogs(string a_sName, string a_sLogPath)
        {
            string sFuncName = "[SearchTestLogs]";

            string a_sStartPath = a_sLogPath + a_sName + "\\";

            //string sTempListForLog = a_sStartPath + a_sName + "_testlogs.list";
            string sTempListFileName = a_sName + "_testlogs.txt";


            // Check the input parameters.
            if ("" == a_sStartPath)
            {
                Logger.Print(sFuncName, "Null start path is specified.");
                return null;
            }

            if (!Directory.Exists(a_sLogPath))
            {
                Logger.Print(sFuncName, $"Can't found direcoty {a_sLogPath}");
                return null;
            }


            List<TestLog> a_lsOutList = new List<TestLog>();


            // Serach and collect all log files recursively.
            CollectFiles(a_sStartPath, Constants.TESTLOG_FILENAME_EXT, ref a_lsOutList);

            // Save the list of found files to the specifed file.
            if ("" != sTempListFileName)
            {
#if DEBUG
                WriteTestLogsListToTextFile(ref a_lsOutList, sTempListFileName);
#endif
            }

            //List<string> logs = processLogs(a_lsOutList, sTempListFileName);

            Logger.Print(sFuncName, a_lsOutList.Count.ToString() + " Test Log(s) collected.");


            return a_lsOutList;
        }
        
        /// <summary>
        /// Collect the test log from file system.
        /// </summary>
        /// <param name="a_sDir">starting folder</param>
        /// <param name="a_sFileExt">extention file name</param>
        /// <param name="a_lsCollection">a List object for TestLog</param>
        /// <returns></returns>
        private List<TestLog> CollectFiles(string a_sDir, string a_sFileExt, ref List<TestLog> a_lsCollection)
        {
            string sFuncName = "[CollectFiles - Test Log]";

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
                    TestLog t = new TestLog(f);
                    a_lsCollection.Add(t);
                }

                // Collect the considered files stored in sub-folders.
                string[] DirList = Directory.GetDirectories(a_sDir);
                foreach (string d in DirList)
                {
                    a_lsCollection = CollectFiles(d, a_sFileExt, ref a_lsCollection);
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
        private bool WriteTestLogsListToTextFile(ref List<TestLog> a_lsInList, string a_sOutFile)
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

            if (File.Exists(a_sOutFile))
            {
                File.Delete(a_sOutFile);
            }

            // Write the error log to the output file.
            try
            {

                using (StreamWriter sw = File.AppendText(a_sOutFile))
                {
                    foreach (TestLog sLine in a_lsInList)
                    {
                        string s1 = sLine.ClassName;
                        string s2 = sLine.FileName;
                        string s3 = sLine.FullPath;

                        string data = $"{s1}, {s2}, {s3}";

                        sw.WriteLine(data);
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


    }
}
