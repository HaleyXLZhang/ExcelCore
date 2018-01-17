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
                excelCore.OpenExcel(@"C:\Users\Administrator\Desktop\TestExcel\Letter Automation V1.19.1.xlsm", false);
                excelCore.SelectSheet("Tool");

                excelCore.SetCellValue(2, "B", "");
                excelCore.SetCellValue(3, "B", "");
                excelCore.SetCellValue(4, "B", "");
                excelCore.SetCellValue(5, "B", "");
                excelCore.SetCellValue(6, "B", "");
                excelCore.SetCellValue(7, "B", "");
                excelCore.SetCellValue(8, "B", "");
                excelCore.SetCellValue(9, "B", "");
                excelCore.SetCellValue(10, "B", "");
                excelCore.SetCellValue(11, "B", "");
                excelCore.SetCellValue(12, "B", "");

                excelCore.Save();

                excelCore.SetCellValue(2, "B", "345");
                excelCore.SetCellValue(3, "B", "345");
                excelCore.SetCellValue(4, "B", "3425");
                excelCore.SetCellValue(5, "B", "345");
                excelCore.SetCellValue(6, "B", "PAH");
                excelCore.SetCellValue(7, "B", "EE");
                excelCore.SetCellValue(8, "B", "PW");
                excelCore.SetCellValue(9, "B", "ER");
                excelCore.SetCellValue(10, "B", "");
                excelCore.SetCellValue(11, "B", "1st letter");
                excelCore.SetCellValue(12, "B", "Y");

                excelCore.Save();

                excelCore.SelectSheet("Data");


                excelCore.SetCellValue(2, "P", "C");


                Cell input_result = excelCore.GetCell(2, "S");
                Cell Match_document_Name_result = excelCore.GetCell(5,"S");
                Cell Match_Reason_document_result = excelCore.GetCell(14, "S");






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
            Microsoft.Office.Interop.Excel.Workbook wbk = app.Workbooks.Open(@"C:\Users\Administrator\Desktop\TestExcel\Letter Automation V1.19.1.xlsm",
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
            Worksheet sheet = (Worksheet)wbk.Worksheets["Data"];
            
            object DeptComboBox=sheet.Range["DeptComboBox"].Value;
            object Language = sheet.Range["Language"].Value;
            object value = sheet.Range["K40", "K46"].Value;
            object language = sheet.Range["P2"].Value;
            object reasonDocName = sheet.Range["S14"].Value;
            sheet.Range["P2"].Value = "C";
            reasonDocName = sheet.Range["S14"].Value;
            object furmula = sheet.Range["S14"].Formula;
        }
    }
}