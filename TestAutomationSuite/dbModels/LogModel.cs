using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TestAutomationSuite.dbModels
{
    public class LogModel
    {
        public int logId { get; set; }

        public int recordId { get; set; }

        public string logTime { get; set; }

        public string description { get; set; }
    }
}
