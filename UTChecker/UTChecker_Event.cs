using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace UTChecker
{


    /// <summary>
    /// ReportProgressEventArgs
    /// </summary>
    public class UTCheckerEvent : EventArgs
    {
        public string Message { get; set; } = string.Empty;
        public int ReturnCode { get; set; } = 0;

        public EnvrionmentSetting Path { get; set; }

        public RunMode Mode { get; set; }
    }



    public partial class UTChecker
    {

        /// <summary>
        /// An event for update 
        /// </summary>
        public event EventHandler<UTCheckerEvent> UpdatePathEvent;


        /// <summary>
        /// 
        /// </summary>
        public event EventHandler<UTCheckerEvent> CompletedEvent;



        /// <summary>
        /// 
        /// </summary>
        /// <param name="e"></param>
        protected virtual void OnUpdatePathEvent(UTCheckerEvent e)
        {
            EventHandler<UTCheckerEvent> handler = UpdatePathEvent;

            if (handler != null)
            {
                handler(this, e);
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="e"></param>
        protected virtual void OnCompletedEvent(UTCheckerEvent e)
        {
            EventHandler<UTCheckerEvent> handler = CompletedEvent;

            if (handler != null)
            {
                handler(this, e);
            }
        }


    }
}
