using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TestAutomationSuite.dbModels
{
    public class ProcessModel
    {
        public int processId { get; set; }

        public int winProcID { get; set; }

        public string startTime { get; set; }

        public string endTime { get; set; }

        public string processStatus { get; set; }
    }
}
