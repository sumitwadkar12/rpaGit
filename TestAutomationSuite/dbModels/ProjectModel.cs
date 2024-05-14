using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TestAutomationSuite.dbModels
{
    public class ProjectModel
    {
        public int projectId { get; set; }
        public string projectName { get; set; }
        public string actionFileJson { get; set; }
        public string inputFilePath { get; set; }
        public string projectStatus {  get; set; }
    }
}
