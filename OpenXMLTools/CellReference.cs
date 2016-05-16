using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXMLTools
{
    class CellReference
    {
        public int RowIndex { get; private set; }
        public int ColumnIndex { get; private set; }
        public string ColumnName { get { return _columnName; } }
        public string CellName { get { return _cellName; } }
        

        public CellReference(int rowIndex, int columnIndex)
        {
            RowIndex = rowIndex;
            ColumnIndex = columnIndex;             
            _columnName = columnList[columnIndex - 1];
            _cellName = _columnName + rowIndex;
        }


        private string _columnName;
        private string _cellName;
        readonly static List<String> columnList = new List<string>("ABCDEFGHIJKLMNOPQRSTUVWXYZ".ToList().Select(c => c.ToString()).ToList());
    }
}
