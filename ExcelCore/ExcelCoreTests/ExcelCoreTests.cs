using Microsoft.VisualStudio.TestTools.UnitTesting;
using ExcelCore;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelCore.Common;

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
                excelCore.OpenExcel(@"C:\Users\li\Desktop\Test.xlsx",false);
                excelCore.SelectSheet("Sheet1");
              Column column=  excelCore.GetColumn("A");
            }
        }
    }
}