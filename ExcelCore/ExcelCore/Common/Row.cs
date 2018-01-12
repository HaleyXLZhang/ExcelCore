using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelCore.Common
{
    public class Row
    {
        public int Index { get; set; }

        public List<Cell> Cells { get; set; }

        public Row(){
            Cells = new List<Cell>();
        }
    }
}
