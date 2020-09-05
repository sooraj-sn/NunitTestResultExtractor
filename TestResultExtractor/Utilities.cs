#region SystemNamespaces
using Microsoft.Office.Interop.Excel;
using System.Collections.Generic;

#endregion

namespace TestResultExtractor
{
    /// <summary>
    /// This class holds global data for this application
    /// </summary>
    public class Utilities
    {
        Application excelApp = new Microsoft.Office.Interop.Excel.Application();

        #region Private Properties
        Application excel;
        Workbook excelworkBook;
        Worksheet excelSheet;
        Range excelCellrange;

        #endregion

        #region FileIO Methods

        /// <summary>
        /// 
        /// </summary>
        /// <param name="accounts"></param>
        public void DisplayInExcel(List<TestResults> testResults,List<Summary> summaryResults)
        {

            // Make the object visible.
            excelApp.Visible = true;

            // Create a new, empty workbook and add it to the collection returned 
            // by property Workbooks. The new workbook becomes the active workbook.
            // Add has an optional parameter for specifying a praticular template. 
            // Because no argument is sent in this example, Add creates a new workbook. 
            excelApp.Workbooks.Add();

            // This example uses a single workSheet. The explicit type casting is
            // removed in a later procedure.
            Microsoft.Office.Interop.Excel._Worksheet workSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelApp.ActiveSheet;
            workSheet.Name = "Details";
            // Establish column headings in cells A1 and B1.
            workSheet.Cells[1, "A"] = "Serial No :";
            workSheet.Cells[1, "B"] = "TestCaseName";
            workSheet.Cells[1, "C"] = "Executed";
            workSheet.Cells[1, "D"] = "Result";
            workSheet.Cells[1, "E"] = "Time";
            workSheet.Cells[1, "F"] = "FailedReason";
            workSheet.Range[workSheet.Cells[1, 1], workSheet.Cells[1, 6]].Font.Bold = true;   
           
            

            var row = 1;
            foreach (var testresult in testResults)
            {
                row++;
                if (testresult.failureReason != "")
                {
                    workSheet.Cells[row, "B"].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                    workSheet.Cells[row, "D"].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                    workSheet.Cells[row,"F"].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);                  
                }
                workSheet.Cells[row, "A"] = testresult.testCount;
                workSheet.Cells[row, "B"] = testresult.testCaseName;
                workSheet.Cells[row, "C"] = testresult.executed;
                workSheet.Cells[row, "D"] = testresult.result;
                workSheet.Cells[row, "E"] = testresult.time;
                workSheet.Cells[row, "F"] = testresult.failureReason;
            }
            workSheet.Columns[1].AutoFit();
            workSheet.Columns[2].AutoFit();
            workSheet.Columns[3].AutoFit();
            workSheet.Columns[4].AutoFit();
            workSheet.Columns[5].AutoFit();
            workSheet.Columns[6].AutoFit();

            excelApp.Worksheets.Add();
            workSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelApp.ActiveSheet;
            workSheet.Name = "Summary";
            workSheet.Range[workSheet.Cells[1, 1], workSheet.Cells[1, 4]].Merge();
            workSheet.Cells[1, "A"] = "Summary";
            workSheet.Range[workSheet.Cells[1, 1], workSheet.Cells[1, 1]].Font.Bold = true;
            workSheet.Cells[2, "A"] = "TotalTestCases";
            workSheet.Cells[3, "A"] = "TotalExecuted";
            workSheet.Cells[4, "A"] = "Total Passed";
            workSheet.Cells[5, "A"] = "Total Failed";

            XmlToCSVEngine xmlEngine = new XmlToCSVEngine();
            row = 2;
            foreach (var summaryresult in summaryResults)
            {
                workSheet.Cells[row, "B"] = summaryresult.totalTestCases;
                row++;
                workSheet.Cells[row, "B"] = summaryresult.totalExecuted;
                row++;
                workSheet.Cells[row, "B"] = summaryresult.totalPassed;
                row++;
                workSheet.Cells[row, "B"] = summaryresult.totalFailed;
                row++;

            }

            workSheet.Range[workSheet.Cells[2, 1], workSheet.Cells[2, 2]].Font.Bold = true;
            workSheet.Range[workSheet.Cells[3, 1], workSheet.Cells[3, 2]].Font.Bold = true;
            workSheet.Range[workSheet.Cells[4, 1], workSheet.Cells[4, 2]].Font.Bold = true;
            workSheet.Range[workSheet.Cells[5, 1], workSheet.Cells[5, 2]].Font.Bold = true;
            workSheet.Columns[1].AutoFit();
            workSheet.Columns[2].AutoFit();
        }     
            #endregion
        
    }
}
