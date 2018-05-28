using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace UTChecker
{
    public partial class UTChecker
    {
        private static int _gLogIndex = -1;
        private static Object LogLock = new object();


        public static int gLogIndex
        {

            get
            {
                lock(LogLock)
                {
                    return _gLogIndex;
                }
            }

            set
            {
                lock (LogLock)
                {
                    _gLogIndex = value;
                }
            }
    
        }

        public static void SearchTestLog(object obj)
        {

            Tuple<string, string, List<TestLog>> items = (Tuple<string, string, List<TestLog>>)obj;

            string className = items.Item1;
            string fileName = items.Item2;
            List<TestLog> a_lsTestLogs = items.Item3;

            Predicate<TestLog> FindValue = delegate (TestLog log)
            {
                return (log.ClassName == className) && (log.FileName == fileName);
            };


            int index = a_lsTestLogs.FindIndex(FindValue);

            gLogIndex = index;
        }



        private delegate int SearchTestLogExDelegate(object obj);

        public static int  SearchTestLogEx(object obj)
        {

            Tuple<string, string, List<TestLog>> items = (Tuple<string, string, List<TestLog>>)obj;

            string className = items.Item1;
            string fileName = items.Item2;
            List<TestLog> a_lsTestLogs = items.Item3;

            Predicate<TestLog> FindValue = delegate (TestLog log)
            {
                return (log.ClassName == className) && (log.FileName == fileName);
            };


            int index = a_lsTestLogs.FindIndex(FindValue);

            gLogIndex = index;

            return index;
        }






        /// <summary>
        /// 
        /// </summary>
        public class TestLog : IComparable<TestLog>
        {

            /// <summary>
            /// 
            /// </summary>
            public enum TestResult
            {
                // Test Log is not available.
                NOT_AVAILABLE = 0,

                // Test Log is Passed
                PASSED,

                // Test Log is Failed
                FAILED,

                // Test Log is Invalid.
                INVALID,
            }


            /// <summary>
            /// The test log are created by which platform.
            /// </summary>
            private enum PlateformType
            {
                /// <summary>
                /// Created by Mockito
                /// </summary>
                Mockito = 1,

                /// <summary>
                /// Created by PowerMockito
                /// </summary>
                PowerMockito = 2,

                /// <summary>
                /// Created VectorCAst
                /// </summary>
                VectorCast = 4,

                /// <summary>
                /// Does not belong to any platform.
                /// </summary>
                None = 8,
            }



            /// <summary>
            /// The Target class that is to be tested.
            /// </summary>
            public string ClassName { get; private set; }

            /// <summary>
            /// The full name of Test Log. (with Ext-FileName .TxT)
            /// </summary>
            public string FileName { get; private set; }


            /// <summary>
            /// The full path of Test Log.
            /// </summary>
            public string FullPath { get; private set; }


            /// <summary>
            /// 
            /// </summary>
            private PlateformType Type = PlateformType.None;


            private const string DIR_VECTORCAST = "vectorcast";
            private const string DIR_POWERMOCKITO = "PowerMockito";

            


            /// <summary>
            /// The Count of Used
            /// </summary>
            public int UsedCount { get; private set; }


            /// <summary>
            /// Constructor for test log
            /// </summary>
            public TestLog()
            {
                this.ClassName = "N/A";
                this.FileName = "N/A";
                this.FullPath = "N/A";
                this.Type = PlateformType.None;
                this.UsedCount = 0;
            }



            /// <summary>
            /// Constructor for test log
            /// </summary>
            /// <param name="path">A path string of test log</param>
            public TestLog(string path)
            {
                if (path != "")
                {

                    this.FullPath = path;

                    // get the file name of test log.
                    this.FileName = Path.GetFileName(path);

                    // get the full path of log file
                    string subDir = Path.GetDirectoryName(path);

                    // get the class from previous direcotry.
                    string parentDir = Directory.GetParent(subDir).FullName;


                    if (subDir.ToLower().EndsWith(DIR_VECTORCAST))
                    {
                        string grandParent = Directory.GetParent(parentDir).FullName;


                        this.ClassName = parentDir.Substring(grandParent.Length + 1); // with '\' plus 1

                        this.Type = PlateformType.VectorCast;

                    }
                    else 
                    {
                        this.ClassName = subDir.Substring(parentDir.Length + 1); // with '\' plus 1

                        if (subDir.Contains(DIR_POWERMOCKITO))
                        {
                            this.Type = PlateformType.PowerMockito;
                        }
                        else
                        {
                            this.Type = PlateformType.Mockito;
                        }
    
                    }

                    // set us
                    this.UsedCount = 0;
                }
                else
                {

                    this.ClassName = "N/A";
                    this.FileName = "N/A";
                    this.FullPath = "N/A";
                    this.Type = PlateformType.None;
                    this.UsedCount = 0;

                }
 
            }


            /// <summary>
            /// 
            /// </summary>
            /// <param name="other"></param>
            /// <returns></returns>
            public int CompareTo(TestLog other)
            {
                string c1 = "N/A";
                string c2 = "A/N";

                 c1 = this.ClassName + "." + this.FileName;
                 c2 = other.ClassName + "." + other.FileName;



                return c1.CompareTo(c2);

            }

            /// <summary>
            /// 
            /// </summary>
            /// <param name="obj"></param>
            /// <returns></returns>
            public override bool Equals(object obj)
            {
                TestLog other = obj as TestLog;

                string c1 = "N/A";
                string c2 = "A/N";

                if (other == null)
                {
                    return false;
                }


                c1 = this.ClassName + "." + this.FileName;
                c2 = other.ClassName + "." + other.FileName;


                if (c1.CompareTo(obj) == 0)
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }

            /// <summary>
            /// 
            /// </summary>
            /// <returns></returns>
            public override int GetHashCode()
            {
                return base.GetHashCode();
            }





            /// <summary>
            /// 
            /// </summary>
            /// <returns></returns>
            public override string ToString()
            {

                if (this.Type == PlateformType.Mockito ||
                    this.Type == PlateformType.PowerMockito)
                {
                    // testXXXTCY.TXT
                    return this.ClassName + "." + this.FileName.Replace(".txt", "");

                }
                else if (this.Type == PlateformType.VectorCast)
                {
                    // XXX.TCYYY.TXT
                    return this.FileName.Replace(".txt", "");
                }
                else
                {
                    return "";
                }
            }


            /// <summary>
            /// 
            /// </summary>
            public void Increment()
            {
                if (this.Type != PlateformType.None)
                {
                    this.UsedCount = this.UsedCount + 1;
                }
            }





            /// <summary>
            /// 
            /// </summary>
            /// <returns></returns>
            public TestResult GetTestResult()
            {
                //string sFuncName = "[GetTestResult]";
                TestResult eTestResult = TestResult.NOT_AVAILABLE;

                

                // Check the existence of current log file.
                if (!File.Exists(this.FullPath))
                {
                    //Log(sFuncName, "Cannot find " + a_sLogFile);
                    return TestResult.NOT_AVAILABLE;
                }

                
                switch (this.Type)
                {

                    case PlateformType.Mockito:
                        eTestResult = ParseMockitoTestResult();
                        break;

                    case PlateformType.PowerMockito:
                        eTestResult = ParsePowerMockitoTestResult();
                        break;

                    case PlateformType.VectorCast:
                        eTestResult = ParseVectorTestResult();
                        break;

                    default:
                        eTestResult = TestResult.NOT_AVAILABLE;
                        break;
                }



                // If passed and failed cases cannot be found, it is an error case.
                return eTestResult;
                
            }


            /// <summary>
            /// 
            /// </summary>
            /// <returns></returns>
            public TestResult GetTestResultByAnalysisRecord(string path)
            {
                return TestResult.NOT_AVAILABLE;
            }



            /// <summary>
            /// Parse test log that created by Mockito
            /// </summary>
            /// <returns></returns>
            private TestResult ParseMockitoTestResult()
            {
                string sFuncName = "[ParseMockitoTestResult]";

                TestResult eTestResult = TestResult.NOT_AVAILABLE;


                if (this.Type == PlateformType.Mockito)
                {

                    string[] sTokens = {"): passed:",
                                        "): failed:"};

                    try
                    {

                        // read all from file.
                        string[] sLines = File.ReadAllLines(FullPath);

                        // remove ext file name.
                        string testCaseString = this.FileName.Replace(".txt", "");


                        //// confirm
                        //if (!sLines[0].Contains(testCaseString) &&
                        //    !sLines[sLines.Length - 1].Contains(testCaseString))
                        //{
                        //    return TestResult.INVALID;
                        //}


                        // Search for the "passed" token from last 5 lines of the file.
                        for (int i = sLines.Length - 1; i >= 0; i--)
                        {
                            if (sLines[i].Contains(sTokens[0]))
                            {
                                eTestResult = TestResult.PASSED;
                                break;
                            }
                            else if (sLines[i].Contains(sTokens[1]))
                            {
                                eTestResult = TestResult.FAILED;
                                break;
                            }
                            else
                            {
                                eTestResult = TestResult.NOT_AVAILABLE;
                            }
                        }

                    }
                    catch (Exception ex)
                    {
                        Logger.Print(sFuncName, ex.Message);
                        eTestResult = TestResult.NOT_AVAILABLE;
                    }
                }
                else
                {
                    eTestResult = TestResult.NOT_AVAILABLE;
                }


                return eTestResult;

            }



            /// <summary>
            /// Parse test log that created by PowerMockito
            /// </summary>
            /// <returns></returns>
            private TestResult ParsePowerMockitoTestResult()
            {
                string sFuncName = "[ParseMockitoTestResult]";

                TestResult eTestResult = TestResult.NOT_AVAILABLE;


                if (this.Type == PlateformType.PowerMockito)
                {

                    string[] sTokens = {"Test Result: Passed",
                                        "Test Result: Failed"};

                    try
                    {

                        // read all from file.
                        string[] sLines = File.ReadAllLines(FullPath);

                        // remove ext file name.
                        string testCaseString = this.FileName.Replace(".txt", "");

                       
                        // confirm
                        //if (!sLines[0].Contains(testCaseString) &&
                        //    !sLines[sLines.Length - 1].Contains(testCaseString))
                        //if (!sLines[sLines.Length - 2].Contains(testCaseString))
                        //    {
                        //    return TestResult.INVALID;
                        //}


                        // Search for the "passed" token from last 5 lines of the file.
                        for (int i = sLines.Length - 1; i >= 0; i--)
                        {
                            if (sLines[i].Contains(sTokens[0]))
                            {
                                eTestResult = TestResult.PASSED;
                                break;
                            }
                            else if (sLines[i].Contains(sTokens[1]))
                            {
                                eTestResult = TestResult.FAILED;
                                break;
                            }
                            else
                            {
                                eTestResult = TestResult.NOT_AVAILABLE;
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Logger.Print(sFuncName, ex.Message);
                        eTestResult = TestResult.NOT_AVAILABLE;

                    }
                }
                else
                {
                    eTestResult = TestResult.NOT_AVAILABLE;
                }

                return eTestResult;

            }





            /// <summary>
            /// Parse test log that created by VectorCast
            /// </summary>
            /// <returns></returns>
            private TestResult ParseVectorTestResult()
            {
                string sFuncName = "[ParseVectorTestResult]";

                TestResult eTestResult = TestResult.NOT_AVAILABLE;


                if (this.Type == PlateformType.VectorCast)
                {

                    const string RESULT_LINE_TOKEN = "Test Status";
                    const string PASS_TOKEN = "PASS";
                    const string FAIL_TOKEN = "FAIL";

                    try
                    {

                        string line;

                        using (System.IO.StreamReader file = new System.IO.StreamReader(FullPath))
                        {
                        
                            while ((line = file.ReadLine()) != null)
                            {
                                line = line.Trim();

                                if (line.StartsWith(RESULT_LINE_TOKEN))
                                {
                                    string result = line.Remove(0, RESULT_LINE_TOKEN.Length).Trim();

                                    if (result.Equals(PASS_TOKEN))
                                    {
                                        eTestResult = TestResult.PASSED;
                                    }
                                    else if (result.Equals(FAIL_TOKEN))
                                    {
                                        eTestResult = TestResult.FAILED;
                                    }
                                    else
                                    {
                                        eTestResult = TestResult.INVALID;
                                    }

                                    break;
                                }
                                else
                                {
                                    eTestResult = TestResult.NOT_AVAILABLE;
                                }
                            }

                        }

                    }
                    catch (Exception ex)
                    {
                        Logger.Print(sFuncName, ex.Message);
                        eTestResult = TestResult.NOT_AVAILABLE;
                    }
                }
                else
                {
                    eTestResult = TestResult.NOT_AVAILABLE;
                }


                return eTestResult;

            }

        }
    }
}
