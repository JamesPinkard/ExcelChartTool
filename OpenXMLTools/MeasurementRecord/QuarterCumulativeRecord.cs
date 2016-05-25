using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OpenXMLTools
{
    class QuarterCumulativeRecord : CumulativeRecord
    {
        public QuarterCumulativeRecord(CumulativeRecord cumulativeRecord, double averageFlow)
        {
            _cumulativeRecord = cumulativeRecord;
            _averageFlow = averageFlow;
        }

        public override int Week
        {
            get
            {
                return _cumulativeRecord.Week;
            }
        }

        public override List<Cell> GetHeader()
        {
            List<Cell> cellList = _cumulativeRecord.GetHeader();
            cellList.Add(new Cell { CellValue = new CellValue("Quarterly Flow Rate \n (GPM)"), DataType = CellValues.String });
            return cellList;
        }

        public override List<Cell> GetFilledCells()
        {
            List<Cell> cellList = _cumulativeRecord.GetFilledCells();
            cellList.Add(new Cell { CellValue = new CellValue(_averageFlow.ToString()), DataType = CellValues.Number });
            return cellList;
        }

        private CumulativeRecord _cumulativeRecord;
        private double _averageFlow;
    }
}
