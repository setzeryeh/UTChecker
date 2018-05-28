using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace UTChecker
{
    /// <summary>
    /// This class is used to record the debug/improtant message into file or LoggerForm
    /// </summary>
    public static class Logger
    {
        /// <summary>
        /// The name of log file, This file is used to record the log data.
        /// </summary>
        public static string FileName { get; set; }


        /// <summary>
        /// The name of log file, This file is used to record the log data in Whole process.
        /// </summary>
        private static string gProcessLogFile = String.Empty;


        /// <summary>
        /// An Event is used to report the message to LoggerForm.
        /// </summary>
        public static event LoggerForm.ReportMessageEventHandler ReportMessageEvent;

        /// <summary>
        /// An event is used to notify the LoggerForm for clearning the message on window.
        /// </summary>
        public static event LoggerForm.ReportMessageEventHandler ClearMessageEvent;


        /// <summary>
        /// An Event is used to report the progress to LoggerForm.
        /// </summary>
        public static event LoggerForm.ReportProgressEventHandler ReportProgressEvent;


        /// <summary>
        ///  An event is used to notify the LoggerForm for clearning the progress.
        /// </summary>
        public static event LoggerForm.ReportProgressEventHandler ClearProgressEvent;


        /// <summary>
        /// Enable showing the message to Logger.
        /// </summary>
        public static bool Enable { get; set; }



        private static ConcurrentQueue<Tuple<string, PrintOption>> msgQueue = null;


        private static Object thisLock = null;

        private static Thread thread = null;

        /// <summary>
        /// An item indicates which platofrm be used for Logger.
        /// </summary>
        public enum PrintOption
        {
            /// <summary>
            /// Print to File
            /// </summary>
            File = 1,

            /// <summary>
            /// Print to Logger Form
            /// </summary>
            Logger = 2,

            /// <summary>
            /// Print to File and Logger both.
            /// </summary>
            Both = 3,
        }




        /// <summary>
        /// static constructor
        /// </summary>
        static Logger()
        {
            CultureInfo culture = new CultureInfo("en-US");


            Enable = true;

            // init fileName
            FileName = String.Empty;
            gProcessLogFile = "Process.log";


            thisLock = new object();
            msgQueue = new ConcurrentQueue<Tuple<string, PrintOption>>();

            thread = new Thread(printThred);
            thread.Name = "Print Thread";
            thread.Start();


            if (File.Exists(gProcessLogFile))
            {
                File.Delete(gProcessLogFile);
            }


            // Log the message to process log
            using (StreamWriter sw = File.AppendText(gProcessLogFile))
            {
                sw.WriteLine("Create Process.log at " + DateTime.Now.ToString(culture));
            }



        }


        public static void PrepareProcessLog(string logPath, string fileName = "Process.log")
        {
            string _path = logPath;


            // Ensure each path is ended with a '\\'.
            if (!_path.EndsWith("\\"))
            {
                _path = _path + "\\";
            }

            string sourceFileName = gProcessLogFile;
            string destFileName = _path + gProcessLogFile;


            if (File.Exists(destFileName))
            {
                File.Delete(destFileName);
            }

            if (File.Exists(sourceFileName))
            {
                File.Move(sourceFileName, destFileName);
            }

            gProcessLogFile = destFileName;

        }


        /// <summary>
        /// 
        /// </summary>
        private static void printThred()
        {
            while(true)
            {
                lock (thisLock)
                {
                    if (msgQueue.IsEmpty)
                    {
                        SpinWait.SpinUntil(() => true, 50);
                       
                    }
                    else
                    {

                        Tuple<string, PrintOption> obj;

                        while (msgQueue.TryDequeue(out obj))
                        {
                            string sMsg = obj.Item1;
                            PrintOption opt = obj.Item2;
                            //string sMsg = $"{a_sFuncName} {a_sMessage}";

                            if (opt == PrintOption.Logger || opt == PrintOption.Both)
                            {
                                ReportMessage(sMsg);
                            }


                            if (opt == PrintOption.File || opt == PrintOption.Both)
                            {

                                try
                                {
                                    if (FileName != "")
                                    {
                                        // Log the message to file.
                                        using (StreamWriter sw = File.AppendText(FileName))
                                        {
                                            sw.WriteLine(sMsg);
                                        }
                                    }

                                }
                                catch (Exception ex)
                                {
                                    sMsg = $"{sMsg} : Exception: {ex.Message}";

                                }
                                finally
                                {

                                    //Log the message to file
                                    using (StreamWriter sw = File.AppendText(gProcessLogFile))
                                    {
                                        sw.WriteLine(sMsg);
                                    }

                                }
                            }
                        }

                    }
                }

            }
        }





        /// <summary>
        /// Print the message to Logger, File or Both (Default is Logger)
        /// </summary>
        /// <param name="a_sMessage"></param>
        /// <param name="opt">PrintOptin, default is PrintOption.Logger</param>
        public static void Print(string a_sMessage, PrintOption opt = PrintOption.Logger)
        {
            lock (thisLock)
            {
                Tuple<string, PrintOption> obj = new Tuple<string, PrintOption>(a_sMessage, opt);

                msgQueue.Enqueue(obj);

                //Print(a_sMessage, "", opt);
            }
        }


        /// <summary>
        /// Print the message to Logger, File or Both (Default is File)
        /// </summary>
        /// <param name="a_sFuncName">The message of function name.</param>
        /// <param name="a_sMessage">The body of message</param>
        /// <param name="opt">PrintOptin, default is PrintOption.File</param>
        public static void Print(string a_sFuncName, string a_sMessage, PrintOption opt = PrintOption.File)
        {
            lock (thisLock)
            {
                string sMsg = $"{a_sFuncName} {a_sMessage}";

                Tuple<string, PrintOption> obj = new Tuple<string, PrintOption>(sMsg, opt);

                msgQueue.Enqueue(obj);

                //if (opt == PrintOption.Logger || opt == PrintOption.Both)
                //{
                //    ReportMessage(sMsg);
                //}



                //if (opt == PrintOption.File || opt == PrintOption.Both)
                //{

                //    try
                //    {
                //        if (FileName != "")
                //        {
                //            // Log the message to file.
                //            using (StreamWriter sw = File.AppendText(FileName))
                //            {
                //                sw.WriteLine(sMsg);
                //            }
                //        }

                //    }
                //    catch (Exception ex)
                //    {
                //        sMsg = $"{sMsg} : Exception: {ex.Message}";

                //    }
                //    finally
                //    {

                //        //Log the message to file
                //        using (StreamWriter sw = File.AppendText(gProcessLogFile))
                //        {
                //            sw.WriteLine(sMsg);
                //        }

                //    }
                //}
            }
        }




        /// <summary>
        /// Update the progress to LoggerForm
        /// </summary>
        /// <param name="value"></param>
        public static void UpdateProgress(int value)
        {
            ReportProgress(value);
        }


        /// <summary>
        /// this method to clear the Message and Progream in LoggerForm
        /// </summary>
        public static void ClearAll()
        {
            _ClearMessage();
            _ClearProgress();
        }


        /// <summary>
        /// 
        /// </summary>
        public static void ClearMessage()
        {
            _ClearMessage();
        }

        /// <summary>
        /// 
        /// </summary>
        public static void ClearProgress()
        {
            _ClearProgress();
        }



        /// <summary>
        /// This method triggers an event to report the message
        /// </summary>
        /// <param name="msg"></param>
        private static void ReportMessage(string msg)
        {
            if (!Enable)
                return;

            if (ReportMessageEvent != null)
            {
                LoggerForm.ReportMessageEventArgs e = new LoggerForm.ReportMessageEventArgs();
                e.message = msg;
                ReportMessageEvent(typeof(Logger), e);
            }

        }


        /// <summary>
        /// This method triggers an event to report the progress
        /// </summary>
        private static void ReportProgress(int progress)
        {
            if (!Enable)
                return;

            if (ReportProgressEvent != null)
            {
                LoggerForm.ReportProgressEventArgs e = new LoggerForm.ReportProgressEventArgs();
                e.progress = progress;
                ReportProgressEvent(typeof(Logger), e);
            }
        }

        /// <summary>
        /// This method triggers an event to clear the message on LoggerForm
        /// </summary>
        private static void _ClearMessage()
        {
            if (!Enable)
                return;

            if (ClearMessageEvent != null)
            {
                ClearMessageEvent(typeof(Logger), new LoggerForm.ReportMessageEventArgs());
            }
        }

        /// <summary>
        /// This method triggers an event to clear the progress on LoggerForm
        /// </summary>
        private static void _ClearProgress()
        {
            if (!Enable)
                return;

            if (ClearProgressEvent != null)
            {
                ClearProgressEvent(typeof(Logger), new LoggerForm.ReportProgressEventArgs());
            }
        }

    }
}
