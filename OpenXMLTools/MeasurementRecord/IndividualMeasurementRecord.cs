using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml;

namespace OpenXMLTools
{
    public class IndividualMeasurementRecord : MeasurementRecord
    {
        public IndividualMeasurementRecord(string stationName, DateTime measureTime, int totalizerReading, double cumulativeTime, int cumulativeFlow)
        {
            StationName = stationName;
            MeasureTime = measureTime;
            TotalizerReading = totalizerReading;
            CumulativeTime = cumulativeTime;
            CumulativeFlow = cumulativeFlow;
        }
        public string StationName { get; set; }
        public DateTime MeasureTime { get; set; }
        public int TotalizerReading { get; set; }
        public double CumulativeTime { get; set; }
        public int CumulativeFlow { get; set; }

        public override List<Cell> GetFilledCells()
        {
            return new List<Cell>()
            {
                new Cell { CellValue = new CellValue(StationName), StyleIndex = (UInt32Value)286U, DataType = CellValues.String},
                new Cell { CellValue = new CellValue(MeasureTime.ToOADate().ToString()), StyleIndex = (UInt32Value)287U},
                new Cell { CellValue = new CellValue(TotalizerReading.ToString()), StyleIndex = (UInt32Value)290U, DataType = CellValues.Number},
                new Cell { CellValue = new CellValue(CumulativeTime.ToString()), StyleIndex = (UInt32Value)285U, DataType = CellValues.Number },
                new Cell { CellValue = new CellValue(CumulativeFlow.ToString()), StyleIndex = (UInt32Value)284U, DataType = CellValues.Number}
            };
        }

        public override List<Cell> GetHeader()
        {
            return new List<Cell>()
            {
                new Cell { CellValue = new CellValue("Station"), StyleIndex = (UInt32Value)297U, DataType = CellValues.String},
                new Cell { CellValue = new CellValue("Date and Time"), StyleIndex = (UInt32Value)297U, DataType = CellValues.String},
                new Cell { CellValue = new CellValue("Total Flow\n (gal)"), StyleIndex = (UInt32Value)297U, DataType = CellValues.String},
                new Cell { CellValue = new CellValue("Cumulative Time\n (min)"), StyleIndex = (UInt32Value)297U, DataType = CellValues.String},
                new Cell { CellValue = new CellValue("Cumulative Flow\n (gal)"), StyleIndex = (UInt32Value)297U, DataType = CellValues.String}
            };
        }
    }
}
