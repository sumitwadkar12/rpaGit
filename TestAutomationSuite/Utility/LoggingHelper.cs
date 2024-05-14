using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TestAutomationSuite.Utility
{
    public class LoggingHelper
    {
        ActionHelper actionHelper = new();
        ErrorLog errorLog = new();
        public LoggingHelper() 
        {
            actionHelper=new ActionHelper();
            errorLog = new ErrorLog();
        }
        public void insertLog(string message)
        {
            //LoggerComponent.Log(TestAutomationSuite.Utility.LogLevel.Information, message);
            Console.WriteLine("[INFORMATION] : " + message);
            errorLog.Write(message);
            actionHelper.WriteToFile(DateTime.Now.ToString("dd-MM-yyyy hh:mm:ss") + message);
        }
    }
}
