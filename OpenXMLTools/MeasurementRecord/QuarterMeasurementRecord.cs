using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml;

namespace OpenXMLTools
{
    public class QuarterMeasurementRecord : MeasurementRecord
    {
        public QuarterMeasurementRecord(WeekMeasurementRecord weekRecord, double averageFlow)
        {
            _weekRecord = weekRecord;
            _averageFlow = averageFlow;
        }

        public override List<Cell> GetFilledCells()
        {
            List<Cell> cellList = _weekRecord.GetFilledCells();
            cellList.Add(new Cell { CellValue = new CellValue(_averageFlow.ToString()), StyleIndex = (UInt32Value)293U, DataType = CellValues.Number });
            return cellList;
        }

        public override List<Cell> GetHeader()
        {
            List<Cell> cellList = _weekRecord.GetHeader();
            cellList.Add(new Cell { CellValue = new CellValue("Quarterly Flow Rate \n (GPM)"), StyleIndex = (UInt32Value)297U, DataType = CellValues.String });
            return cellList;
        }

        private double _averageFlow;
        private WeekMeasurementRecord _weekRecord;
    }
}
