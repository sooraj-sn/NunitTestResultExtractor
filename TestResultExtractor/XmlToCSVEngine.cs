#region SystemNamespaces
using System.Collections.Generic;
using System.Xml.Linq;
using System.Data;
using System.Xml;
#endregion

namespace TestResultExtractor
{
    /// <summary>
    /// XmlToCSVEngine defines methods that prepares test results and calls the ExcelMethods
    /// </summary>
    public class XmlToCSVEngine
    {
        /// <summary>
        /// testResults
        /// </summary>
        private static List<TestResults> testResults = new List<TestResults>();
        private static List<Summary> summaryResults = new List<Summary>();

        /// <summary>
        /// GetTestResultSet 
        /// </summary>
        /// <param name="xmlFileName"></param>
        /// <returns cref="testResults">testResults</returns>
        public List<TestResults> GetTestResultSet(string xmlFileName)
        {
            XmlParser parser = new XmlParser();
            
            XElement doc = XElement.Load(xmlFileName);

            testResults = parser.GetTestResults(doc);
            return testResults;
        }
        public List<Summary> GetSummaryList(string xmlFileName)
        {
            XmlParser parser = new XmlParser();

            XmlDocument doc = new XmlDocument();
            doc.Load(xmlFileName);

            summaryResults = parser.GetSummaryDetails(doc);
            return summaryResults;


        }

        /// <summary>
        /// Displays TestResults In an Excel File
        /// </summary>
        /// <param name="testResults"></param>
        public void DisplayResultsInExcel(List<TestResults> testResults,List<Summary> summaryDetails)
        {
            Utilities util = new Utilities();
            util.DisplayInExcel(testResults,summaryDetails);
            
        }

        
        /// <summary>
        /// GetTestResultsInDataTable
        /// </summary>
        /// <returns></returns>
        public DataTable GetTestResultsInDataTable(List<TestResults> testResults)
        {
            ListToDataTableConverter dtmaker = new ListToDataTableConverter();
            var results = dtmaker.ToDataTable(testResults);
            return results;
        }
    }
}
