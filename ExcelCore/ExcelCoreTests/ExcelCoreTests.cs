using ExcelCore.Common;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;
namespace ExcelCore.Tests
{
    [TestClass()]
    public class ExcelCoreTests
    {
        [TestMethod()]
        public void ExcelCoreTest()
        {
            using (ExcelCore excelCore = new ExcelCore())
            {
                excelCore.OpenExcel(@"C:\Users\Administrator\Desktop\TestExcel\Letter Automation V1.19.1.xlsm", false);
                IList<string> sheetNames = excelCore.GetSheetNames();
                excelCore.SelectSheet("Data");
                //DeptComboBox
                IList<Cell> language = excelCore.GetRangeByName("Language");

                IList<Cell> departmentLevel = excelCore.GetRange(new Cell() { ColumnName = "K", RowIndex = 42 },
                    new Cell() { ColumnName = "K", RowIndex = 46 });

                IList<Cell> clientLevel = excelCore.GetRange(new Cell() { ColumnName = "L", RowIndex = 40 },
                  new Cell() { ColumnName = "L", RowIndex = 41 });

                IList<Cell> EELetterlevel = excelCore.GetRange(new Cell() { ColumnName = "M", RowIndex = 40 },
                  new Cell() { ColumnName = "M", RowIndex = 48 });

                IList<Cell> EETransactionlevel = excelCore.GetRange(new Cell() { ColumnName = "N", RowIndex = 40 },
                new Cell() { ColumnName = "N", RowIndex = 72 });
            }
        }

        [TestMethod]
        public void TestClearToolSheet()
        {
            using (ExcelCore excelCore = new ExcelCore())
            {
                excelCore.OpenExcel(@"C:\Users\Administrator\Desktop\WorkFiles\Letter Automation\Control turnaround.xlsx", false);
                excelCore.SelectSheet("Sheet1");

                IList<Row> rows = excelCore.GetSheetByRow();

            }

        }






        [TestMethod()]
        public void ExcelTest()
        {
            Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.ApplicationClass();
            app.EnableMacroAnimations = false;
            app.DisplayAlerts = false;
            app.ScreenUpdating = false;
            app.EnableEvents = false;
            Microsoft.Office.Interop.Excel.Workbook wbk = app.Workbooks.Open(@"C:\Users\Administrator\Desktop\WorkFiles\Letter Automation\Control turnaround.xlsx",
                  Type.Missing,
                  false,
                  Type.Missing,
                  Type.Missing,
                  Type.Missing,
                  Type.Missing,
                  Type.Missing,
                  Type.Missing,
                  true,
                  Type.Missing,
                  Type.Missing,
                  Type.Missing,
                  Type.Missing,
                  Type.Missing);
            Worksheet sheet = (Worksheet)wbk.Worksheets["Sheet1"];


         


            //object DeptComboBox=sheet.Range["DeptComboBox"].Value;
            //object Language = sheet.Range["Language"].Value;
            //object value = sheet.Range["K40", "K46"].Value;
            //object language = sheet.Range["P2"].Value;
            //object reasonDocName = sheet.Range["S14"].Value;
            //sheet.Range["P2"].Value = "C";
            //reasonDocName = sheet.Range["S14"].Value;
            //object furmula = sheet.Range["S14"].Formula;
        }
    }
}