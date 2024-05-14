using System;
using System.Collections.Generic;
using System.IO.Enumeration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TestAutomationSuite.dbModels
{
    public class FileModels
    {
        public int fileId { get; set; }
        public int processId { get; set; }
        public int projectId { get; set; }
        public string fileName { get; set; }        
        public string startTime { get; set; }
        public string endTime { get; set; }
        public string status { get; set; }
        public string outputfilename {  get; set; }
    }
}
