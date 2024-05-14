using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutomationUAT.ViewModels
{
    public class ActionModel
    {
        public string SrNo { get; set; }
        public string ActionType { get; set; }
        public string Description { get; set; }
        public string InputColumnName { get; set; }
        public string SearchBy { get; set; }
        public string Value { get; set; }
        public string WaitTime { get; set; }
        public string Stage { get; set; }
        public string FileType { get; set; }
        public string FilePath { get; set; }
        public string IfErrorOccurs { get; set; }
        public string CheckElement { get; set; }
        public string ImageFlag { get; set; }
        public string ImageAction { get; set; }
        public string ImageWaitTime { get; set; }        
        public string TimeRequired { get; set; }
        public string InputValue { get; set; }
    }
}
