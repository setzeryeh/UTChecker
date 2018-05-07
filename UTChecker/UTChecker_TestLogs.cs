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
        public class TestLog : IComparable<TestLog>
        {

            /// <summary>
            /// 
            /// </summary>
            public string ClassName { get; private set; }

            /// <summary>
            /// 
            /// </summary>
            public string FileName { get; private set; }


            /// <summary>
            /// 
            /// </summary>
            public string FullPath { get; private set; }


            private const string C_VECTORCAST = "vectorcast";
            private bool IsJava = false;


            /// <summary>
            /// The Count of Used
            /// </summary>
            public int UsedCount { get; private set; }

            public TestLog()
            {
                this.ClassName = "N/A";
                this.FileName = "N/A";
                this.FullPath = "N/A";
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


                    if (subDir.ToLower().EndsWith(C_VECTORCAST))
                    {
                        string grandParent = Directory.GetParent(parentDir).FullName;


                        this.ClassName = parentDir.Substring(grandParent.Length + 1); // with '\' plus 1

                        this.IsJava = false;

                    }
                    else
                    {
                        this.ClassName = subDir.Substring(parentDir.Length + 1); // with '\' plus 1

                        this.IsJava = true;
                    }

                    // set us
                    this.UsedCount = 0;
                }
                else
                {
                    this.ClassName = "N/A";
                    this.FileName = "N/A";
                    this.FullPath = "N/A";

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

                if (IsJava)
                {
                    c1 = this.ClassName + "." + this.FileName.Replace(".txt", "");
                    c2 = other.ClassName + "." + other.FileName.Replace(".txt", "");
                }

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

                if (other != null)
                {
                    return false;
                }

                string c1 = "N/A";
                string c2 = "A/N";

                if (IsJava)
                {
                    c1 = this.ClassName + "." + this.FileName.Replace(".txt", "");
                    c2 = other.ClassName + "." + other.FileName.Replace(".txt", "");
                }


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
                if (IsJava)
                {
                    return this.ClassName + "." + this.FileName.Replace(".txt", "");
                }
                else
                {
                    return this.FileName.Replace(".txt", ""); ;
                }
            }


            /// <summary>
            /// 
            /// </summary>
            public void Increment()
            {
                this.UsedCount = this.UsedCount + 1;
            }

        }




    }
}
