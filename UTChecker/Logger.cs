using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
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
        private static string Process;

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


        public enum PrintOption
        {
            File = 1,
            Logger = 2,
            Both = 3,
        }


        /// <summary>
        /// static constructor
        /// </summary>
        static Logger()
        {
            FileName = "";
            Process = "Process.log";

            if (File.Exists(Process))
            {
                File.Delete(Process);
            }


            // Log the message to process log
            using (StreamWriter sw = File.AppendText(Process))
            {
                sw.WriteLine(DateTime.Now.ToString(new CultureInfo("en-US")));
            }

        }

        /// <summary>
        /// Print the message to Logger, File or Both (Default is Logger)
        /// </summary>
        /// <param name="a_sMessage"></param>
        /// <param name="opt">PrintOptin, default is PrintOption.Logger</param>
        public static void Print(string a_sMessage, PrintOption opt = PrintOption.Logger)
        {
            Print(a_sMessage, "", opt);  
        }


        /// <summary>
        /// Print the message to Logger, File or Both (Default is File)
        /// </summary>
        /// <param name="a_sFuncName">The message of function name.</param>
        /// <param name="a_sMessage">The body of message</param>
        /// <param name="opt">PrintOptin, default is PrintOption.File</param>
        public static void Print(string a_sFuncName, string a_sMessage, PrintOption opt = PrintOption.File)
        {

            string sMsg = $"{a_sFuncName} {a_sMessage}";


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
                    // Log the message to process log
                    using (StreamWriter sw = File.AppendText(Process))
                    {
                        sw.WriteLine(sMsg);
                    }

                }
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
        public static void Clear()
        {
            ClearMessage();
            ClearProgress();
        }





        /// <summary>
        /// This method triggers an event to report the message
        /// </summary>
        /// <param name="msg"></param>
        private static void ReportMessage(string msg)
        {
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
        private static void ClearMessage()
        {
            if (ClearMessageEvent != null)
            {
                ClearMessageEvent(typeof(Logger), new LoggerForm.ReportMessageEventArgs());
            }
        }

        /// <summary>
        /// This method triggers an event to clear the progress on LoggerForm
        /// </summary>
        private static void ClearProgress()
        {
            if (ClearProgressEvent != null)
            {
                ClearProgressEvent(typeof(Logger), new LoggerForm.ReportProgressEventArgs());
            }
        }

    }
}
