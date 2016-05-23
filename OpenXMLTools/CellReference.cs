using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXMLTools
{
    public class CellReference
    {
        public int RowIndex { get; private set; }
        public int ColumnIndex { get; private set; }
        public string ColumnName { get { return GetColumnName(); } }
        public string CellName { get { return GetCellName(); } }

        public CellReference(string cellAddress)
        {
            var columnPart = cellAddress.Substring(0,1);
            var rowPart = cellAddress.Substring(1);
            ColumnIndex = columnList.FindIndex(s => s == columnPart) + 1;
            RowIndex = int.Parse(rowPart);            
        }

        public CellReference(int rowIndex, int columnIndex)
        {
            RowIndex = rowIndex;
            ColumnIndex = columnIndex;             
        }


        private string GetCellName()
        {
            return GetColumnName() + RowIndex;
        }

        private string GetColumnName()
        {
            return columnList[ColumnIndex - 1];
        }

        public void Offset(int rowOffset, int columnOffset)
        {
            RowIndex += rowOffset;
            ColumnIndex += columnOffset;
        }

        public override string ToString()
        {
            return GetCellName();
        }

        readonly static List<String> columnList = new List<string>("ABCDEFGHIJKLMNOPQRSTUVWXYZ".ToList().Select(c => c.ToString()).ToList());
    }
}
