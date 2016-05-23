using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXMLTools
{
    public class CellRange
    {
        public CellRange(string startCellAddress, string endCellAddress)
        {
            _startCell = new CellReference(startCellAddress);
            _endCell = new CellReference(endCellAddress);
        }

        public string GetRangeReference()
        {
            return string.Format("{0}:{1}", _startCell.ToString(), _endCell.ToString());
        }

        public void ExtendRange(int rowOffset, int columnOffset)
        {
            _endCell.Offset(rowOffset, columnOffset);
        }

        public void MoveRange(int rowOffset, int columnOffset)
        {
            _startCell.Offset(rowOffset, columnOffset);
            _endCell.Offset(rowOffset, columnOffset);
        }

        public bool IsValidRange()
        {
            return _startCell.CellName != _endCell.CellName;
        }

        CellReference _startCell;
        CellReference _endCell;
    }
}
