using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TestAutomationSuite.dbModels
{
    public class RecordModel
    {
        public int recordId { get; set; }
        public int processId { get; set; }
        public int fileId { get; set; }
        public string inputJsonRow { get; set; }
        public string startTime { get; set; }
        public string endTime { get; set; }
        public string outputJsonRow { get; set; }
        public string status { get; set; }
        public string lastStageProcessed { get; set; }
        public string lastError { get; set; }

        public string OutputfileName { get; set; }


    }
}
