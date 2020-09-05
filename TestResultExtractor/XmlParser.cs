#region System namespaces
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Linq;
using System.Xml.XPath;
using Microsoft.Office.Interop.Excel;

#endregion

namespace TestResultExtractor
{
    /// <summary>
    /// XmlParser
    /// </summary>
    public class XmlParser
    {
        public List<TestResults> testResults = new List<TestResults>();
        public List<Summary> summaryResults = new List<Summary>();
        public List<TestResults> GetTestResults(XElement doc)
        {
            
            GetParameterisedTestResults(doc);
            return testResults;
        }

        public List<Summary> GetSummaryDetails(XmlDocument doc)
        {
            GetDetailsForSummary(doc);
            return summaryResults;
        }

        /// <summary>
        /// Get Parameterised TestResults
        /// </summary>
        /// <param name="doc"></param>
        /// <returns></returns>
        public void GetParameterisedTestResults(XElement doc)
        {
          
            int testCount = 1;            
            IEnumerable<XElement> testcases =
                from testcase in doc.Descendants("test-case")
                select testcase;
            foreach (XElement testcase in testcases)
            {                
                TestResults parsedresult = new TestResults();
                parsedresult.testCount =testCount++;
                parsedresult.testCaseName = testcase.Attribute("name").Value;
                parsedresult.executed = testcase.Attribute("executed").Value;
                parsedresult.result = testcase.Attribute("result").Value;
                parsedresult.time = testcase.Attribute("time").Value;
                if (testcase.Attribute("result").Value.Equals("Failure") || testcase.Attribute("result").Value.Equals("Error"))
                {
                    parsedresult.failureReason = testcase.LastNode.ToString();

                }
                else
                {
                    parsedresult.failureReason = "";
                }
                testResults.Add(parsedresult);
            }


        }

        public void GetDetailsForSummary(XmlDocument doc)
        {
            XmlNodeList xnList = doc.SelectNodes("test-results");
            foreach (XmlNode xn in xnList)
            {
                Summary summary = new Summary();
                int totalExecuted = Convert.ToInt16(xn.Attributes["total"].Value) - Convert.ToInt16(xn.Attributes["not-run"].Value);
                int totalPassed = Convert.ToInt16(xn.Attributes["total"].Value) - Convert.ToInt16(xn.Attributes["failures"].Value)-Convert.ToInt32(xn.Attributes["errors"].Value);
                int totalFailed= Convert.ToInt16(xn.Attributes["failures"].Value) +Convert.ToInt32(xn.Attributes["errors"].Value);
                summary.totalTestCases= xn.Attributes["total"].Value.ToString();
                summary.totalExecuted = totalExecuted.ToString();
                summary.totalPassed = totalPassed.ToString();
                summary.totalFailed = totalFailed.ToString();

                summaryResults.Add(summary);
            }
                
        }
    }
}
