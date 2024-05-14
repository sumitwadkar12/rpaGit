using System;
using System.Collections.Generic;
using System.Configuration;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TestAutomationSuite.Utility
{
    public class ErrorLog
    {
        Queue<string> qMsg;
        //Thread to write messages
        Thread tLog;
        //to trigger message received
        AutoResetEvent areRestLog;
        //for locking
        string LogDirectory;
        object objToLock;
        string sFilePath;
        public static ErrorLog Instance = new ErrorLog();

        public ErrorLog()
        {
            //for log directory
            LogDirectory = ConfigurationManager.AppSettings["LogDirectory"];
            objToLock = new object();
            qMsg = new Queue<string>();
            areRestLog = new AutoResetEvent(true);
            tLog = new Thread(new ThreadStart(writeMsg));
            tLog.IsBackground = true;
            tLog.Start();
    
            sFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, LogDirectory);

            if (!Directory.Exists(sFilePath))
            {
                sFilePath = System.IO.Directory.GetCurrentDirectory();
            }
        }

        public void Write(string sMsg)
        {
            lock (objToLock)
            {
                qMsg.Enqueue(sMsg);
            }

            areRestLog.Set();
        }

        string dequeueMsg()
        {
            string sMsg;
            lock (objToLock)
            {
                sMsg = qMsg.Dequeue();
            }

            return sMsg;
        }

        void writeMsg()
        {
            while (true)
            {

                //this variable used to create log filename format "
                //for example filename : ErrorLogYYYYMMDD
                string sYear = DateTime.Now.Year.ToString();
                string sMonth = 0 + DateTime.Now.Month.ToString();
                string sDay = 0 + DateTime.Now.Day.ToString();

                string sLogFileNM = "Log" + sYear + sMonth.Substring(sMonth.Length - 2) + sDay.Substring(sDay.Length - 2);

                while (qMsg.Count > 0)
                {
                    string sMsg = dequeueMsg();

                    try
                    {
                        StreamWriter sw = new StreamWriter(sFilePath + @"\" + sLogFileNM, true);
                        sw.WriteLine(DateTime.Now.ToLongTimeString().ToString() + " ==> " + sMsg);
                        sw.Flush();
                        sw.Close();
                    }
                    catch (Exception ex)
                    {
                        EventLog m_EventLog = new EventLog("");
                        m_EventLog.Source = "MySampleEventLog";
                        m_EventLog.WriteEntry("Reading text file failed " + ex.ToString(),EventLogEntryType.FailureAudit);
                    }
                }
                areRestLog.WaitOne();
            }
        }
    }
}
