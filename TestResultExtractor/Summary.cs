using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TestResultExtractor
{
    public class Summary
    {
        public string totalTestCases
        {
            get;set;
        }
        public string totalExecuted
        {
            get; set;
        }
        public string totalPassed
        {
            get; set;
        }
        public string totalFailed
        {
            get; set;
        }
    }
}
