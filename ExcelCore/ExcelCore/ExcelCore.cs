﻿using ExcelCore.Common;
using ExcelCore.Interface;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelCore
{
    public class ExcelCore : IExcel
    {

        internal static dynamic app = null;
        internal dynamic wkb = null;
        public int UsedRowCount = 0;
        public int UsedColumnCount = 0;
        public ExcelCore()
        {
            //设置程序运行语言
            System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");
            app = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application"));
            app.DisplayAlerts = false;
            //设置是否显示Excel
            app.Visible = false;
            //禁止刷新屏幕
            app.ScreenUpdating = false;

        }

        public void Close()
        {
            wkb.Close(Type.Missing, Type.Missing, Type.Missing);
            app.Quit();
            wkb = null;
            app = null;
            GC.Collect();
        }

        public void CreateExcel(string file)
        {
            wkb = app.Workbooks.Add(true);

            wkb.SaveAs(file, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
        }

        public void Dispose()
        {
            Close();
        }

        public Cell GetCell(int rowIndex, string columnName)
        {
            dynamic rng = app.ActiveSheet.get_Range(columnName + rowIndex, columnName + rowIndex);

            object[,] exceldata = (object[,])rng.get_Value(XlRangeValueDataType.xlRangeValueDefault);

            Cell cell = new Cell() { Value = exceldata[rowIndex, ExcelConvert.ToIndex(columnName)+1].ToString(), ColumnName = columnName, RowIndex = rowIndex };

            return cell;
        }

        public IList<Cell> GetRange(Cell start, Cell end)
        {
            List<Cell> cells = new List<Cell>();
            dynamic c1 = app.ActiveSheet.Cells[start.RowIndex, ExcelConvert.ToIndex(start.ColumnName)+1];
            dynamic c2 = app.ActiveSheet.Cells[end.RowIndex, ExcelConvert.ToIndex(end.ColumnName)+1];
            dynamic rng = app.ActiveSheet.get_Range(c1, c2);
            object[,] exceldata = (object[,])rng.get_Value(XlRangeValueDataType.xlRangeValueDefault);
            for (int i = 1; i < exceldata.GetLongLength(0); i++)
            {
                for (int j = 1; j < exceldata.GetLongLength(1); j++)
                {
                    Cell cell = new Cell() { Value = exceldata[i, j].ToString(), ColumnName = ExcelConvert.ToName(j-1), RowIndex = i };
                    cells.Add(cell);
                }
            }
            return cells;
        }

        public Column GetColumn(string columnName)
        {
            Column column = new Column();
            dynamic c1 = app.ActiveSheet.Cells[1, ExcelConvert.ToIndex(columnName)+1];
            dynamic c2 = app.ActiveSheet.Cells[UsedRowCount, ExcelConvert.ToIndex(columnName)+1];
            dynamic rng = app.ActiveSheet.get_Range(c1, c2);
            object[,] exceldata = (object[,])rng.get_Value(XlRangeValueDataType.xlRangeValueDefault);
            for (int i = 1; i < exceldata.GetLongLength(0); i++)
            {
                for (int j = 1; j < exceldata.GetLongLength(1); j++)
                {
                    Cell cell = new Cell() { Value = exceldata[i, j].ToString(), ColumnName = ExcelConvert.ToName(j-1), RowIndex = i };
                    column.Cells.Add(cell);
                }
            }
            return column;
        }

        public IList<Row> GetSheetByRow()
        {
            throw new NotImplementedException();
        }

        public IList<string> GetSheetNames()
        {
            List<string> sheetNames = new List<string>();
            foreach (dynamic sheet in wkb.Worksheets())
            {
                sheetNames.Add(sheet.Name);
            }
            return sheetNames;
        }

        public void OpenExcel(string fileName, bool isReadOnly)
        {
            wkb = app.Workbooks.Open(fileName,
                  Type.Missing,
                  isReadOnly,
                  Type.Missing,
                  Type.Missing,
                  Type.Missing,
                  Type.Missing,
                  Type.Missing,
                  Type.Missing,
                  Type.Missing,
                  Type.Missing,
                  Type.Missing,
                  Type.Missing,
                  Type.Missing,
                  Type.Missing);
            wkb.Activate();
        }

        public void Save()
        {
            wkb.Save();
        }

        public void SaveAs(string fileName)
        {
            wkb.SaveAs(fileName, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
        }

        public void SelectSheet(string sheetName)
        {
            dynamic xlsWorkSheet = wkb.Worksheets[sheetName];
            xlsWorkSheet.Select();
            UsedRowCount = xlsWorkSheet.UsedRange.Rows.Count;
            UsedColumnCount = xlsWorkSheet.UsedRange.Columns.Count;
        }

        public void SelectSheet(int sheetIndex)
        {
            dynamic xlsWorkSheet = wkb.Worksheets(sheetIndex);
            xlsWorkSheet.Select();
            UsedRowCount = xlsWorkSheet.UsedRange.Rows.Count;
            UsedColumnCount = xlsWorkSheet.UsedRange.Columns.Count;
        }

        public void SetCellValue(int rowIndex, string columnName, string value)
        {
            app.ActiveSheet.Cells[rowIndex, ExcelConvert.ToIndex(columnName)+1] = value;
        }
        public void SetCellValue(string sheetName, int rowIndex, string columnName, string value)
        {
            dynamic xlsWorkSheet = wkb.Worksheets[sheetName];
            xlsWorkSheet.Cells[rowIndex, ExcelConvert.ToIndex(columnName)+1] = value;
        }
    }
}