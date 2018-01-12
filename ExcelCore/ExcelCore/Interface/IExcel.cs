using ExcelCore.Common;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelCore.Interface
{
    public interface IExcel : IDisposable
    {
        void CreateExcel(string file);
        void OpenExcel(string fileName, bool isReadOnly);
        void SelectSheet(string sheetName);
        void SelectSheet(int sheetIndex);
        IList<string> GetSheetNames();
        IList<Row> GetSheetByRow();
        Column GetColumn(string columnName);
        IList<Cell> GetRange(Cell start, Cell end);
        void SetCellValue(int rowIndex, string columnName, string value);
        void SetCellValue(string sheetName, int rowIndex, string columnName, string value);
        Cell GetCell(int rowIndex, string columnName);
        void Close();
        void Save();
        void SaveAs(string fileName);
    }
}
