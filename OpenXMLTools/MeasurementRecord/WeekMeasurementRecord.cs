using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml;

namespace OpenXMLTools
{
    public class WeekMeasurementRecord : MeasurementRecord
    {
        public WeekMeasurementRecord(IndividualMeasurementRecord record, int weekIndex, double cumulativeTime, double cumulativeFlow, double avgFlow)
        {
            _record = record;
            _weekIndex = weekIndex;
            _cumulativeTime = cumulativeTime;
            _cumulativeFlow = cumulativeFlow;
            _averageFlow = avgFlow;
        }
        public override List<Cell> GetFilledCells()
        {
            List<Cell> cellList = _record.GetFilledCells();
            cellList.Add(new Cell { CellValue = new CellValue(string.Format("Week {0}", _weekIndex)), StyleIndex = (UInt32Value)294U, DataType = CellValues.String });
            cellList.Add(new Cell { CellValue = new CellValue(_cumulativeTime.ToString()), StyleIndex = (UInt32Value)282U, DataType = CellValues.Number });
            cellList.Add(new Cell { CellValue = new CellValue(_cumulativeFlow.ToString()), StyleIndex = (UInt32Value)282U, DataType = CellValues.Number });
            cellList.Add(new Cell { CellValue = new CellValue(_averageFlow.ToString()), StyleIndex = (UInt32Value)281U, DataType = CellValues.Number });
            return cellList;
        }

        public override List<Cell> GetHeader()
        {
            List<Cell> cellList = _record.GetHeader();
            cellList.Add(new Cell { CellValue = new CellValue("Week #\n"), StyleIndex = (UInt32Value)297U, DataType = CellValues.String });
            cellList.Add(new Cell { CellValue = new CellValue("Weekly Cumulative\n Time"), StyleIndex = (UInt32Value)297U, DataType = CellValues.String });
            cellList.Add(new Cell { CellValue = new CellValue("Weekly Cumulative\n Flow"), StyleIndex = (UInt32Value)297U, DataType = CellValues.String });
            cellList.Add(new Cell { CellValue = new CellValue("Weekly Cumulative\n Flow Rate\n (GPM)"), StyleIndex = (UInt32Value)297U, DataType = CellValues.String });
            return cellList;
        }

        IndividualMeasurementRecord _record;
        private int _weekIndex;
        private double _cumulativeTime;
        private double _cumulativeFlow;
        private double _averageFlow;
    }
}
